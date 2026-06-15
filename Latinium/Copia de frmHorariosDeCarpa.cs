using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmHorariosDeCarpa.
	/// </summary>
	public class frmHorariosDeCarpa : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbVendedores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana2;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana4;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana3;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana5;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCom_Ubicacion1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCom_Ubicacion2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCom_Ubicacion3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCom_Ubicacion4;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCom_Ubicacion5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.DateTimePicker dtpMes;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdHorariosDeCarpa;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Button btnAprobar;
		private System.Windows.Forms.Button btnAprobarSemana1;
		private System.Windows.Forms.Button btnAprobarSemana2;
		private System.Windows.Forms.Button btnAprobarSemana3;
		private System.Windows.Forms.Button btnAprobarSemana4;
		private System.Windows.Forms.Button btnAprobarSemana5;
		private System.Windows.Forms.Button btnAprobarSemana6;
		private System.Windows.Forms.Button btnHorarioCarpa;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmHorariosDeCarpa()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbCom_Ubicacion1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbCom_Ubicacion1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbCom_Ubicacion1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbCom_Ubicacion1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbCom_Ubicacion1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbCom_Ubicacion1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbCom_Ubicacion1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Presupuesto");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Presupuesto", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Presupuesto", 11, true);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Presupuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbCom_Ubicacion2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Presupuesto");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Presupuesto", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Presupuesto", 11, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Presupuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbCom_Ubicacion4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbCom_Ubicacion4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbCom_Ubicacion4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbCom_Ubicacion4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbCom_Ubicacion4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbCom_Ubicacion4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbCom_Ubicacion4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Presupuesto");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Presupuesto", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Presupuesto", 11, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Presupuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbCom_Ubicacion3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbCom_Ubicacion3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbCom_Ubicacion3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbCom_Ubicacion3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbCom_Ubicacion3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbCom_Ubicacion3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbCom_Ubicacion3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Presupuesto");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Presupuesto", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Presupuesto", 11, true);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Presupuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Presupuesto");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Presupuesto", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Presupuesto", 11, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Presupuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Ubicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmHorariosDeCarpa));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Presupuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorariosDeCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbCom_Ubicacion5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Presupuesto");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Presupuesto", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Presupuesto", 11, true);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.uGridSemana1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSemana2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSemana4 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSemana3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSemana5 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtpMes = new System.Windows.Forms.DateTimePicker();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbCom_Ubicacion1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbVendedores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCom_Ubicacion2 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCom_Ubicacion3 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCom_Ubicacion4 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCom_Ubicacion5 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtIdHorariosDeCarpa = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSemana6 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnAprobar = new System.Windows.Forms.Button();
			this.btnAprobarSemana1 = new System.Windows.Forms.Button();
			this.btnAprobarSemana2 = new System.Windows.Forms.Button();
			this.btnAprobarSemana3 = new System.Windows.Forms.Button();
			this.btnAprobarSemana4 = new System.Windows.Forms.Button();
			this.btnAprobarSemana5 = new System.Windows.Forms.Button();
			this.btnAprobarSemana6 = new System.Windows.Forms.Button();
			this.btnHorarioCarpa = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdHorariosDeCarpa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 18);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 727;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(256, 18);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(26, 16);
			this.label4.TabIndex = 725;
			this.label4.Text = "Mes";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 180;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(48, 16);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 21);
			this.cmbBodega.TabIndex = 726;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// uGridSemana1
			// 
			this.uGridSemana1.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 116;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 89;
			ultraGridColumn6.Header.Caption = "Vendedor";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn6.Width = 190;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 46;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 75;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn12.Width = 75;
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn13.Width = 75;
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn14.Width = 75;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.ImageHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance2;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.Header.Caption = "Sector de Carpa";
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn16.Width = 200;
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
																										 ultraGridColumn16});
			ultraGridBand2.Header.Caption = "Semana 1";
			ultraGridBand2.HeaderVisible = true;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance3;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			this.uGridSemana1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana1.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridSemana1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana1.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 7.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana1.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridSemana1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana1.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridSemana1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana1.Location = new System.Drawing.Point(0, 40);
			this.uGridSemana1.Name = "uGridSemana1";
			this.uGridSemana1.Size = new System.Drawing.Size(1024, 104);
			this.uGridSemana1.TabIndex = 730;
			this.uGridSemana1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana1_AfterRowInsert);
			this.uGridSemana1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana1_KeyDown);
			this.uGridSemana1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana1_AfterCellUpdate);
			this.uGridSemana1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSemana1_InitializeLayout);
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
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
																																 ultraDataColumn13});
			// 
			// uGridSemana2
			// 
			this.uGridSemana2.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana2.DataSource = this.ultraDataSource4;
			appearance10.BackColor = System.Drawing.Color.White;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana2.DisplayLayout.Appearance = appearance10;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 116;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 89;
			ultraGridColumn19.Header.Caption = "Vendedor";
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn19.Width = 190;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 46;
			ultraGridColumn21.Header.VisiblePosition = 4;
			ultraGridColumn21.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn21.Width = 75;
			ultraGridColumn22.Header.VisiblePosition = 5;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn22.Width = 75;
			ultraGridColumn23.Header.VisiblePosition = 6;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn23.Width = 75;
			ultraGridColumn24.Header.VisiblePosition = 7;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn24.Width = 75;
			ultraGridColumn25.Header.VisiblePosition = 8;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn25.Width = 75;
			ultraGridColumn26.Header.VisiblePosition = 9;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn26.Width = 75;
			ultraGridColumn27.Header.VisiblePosition = 10;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn27.Width = 75;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance11;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 11;
			ultraGridColumn29.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn29.Header.Caption = "Sector de Carpa";
			ultraGridColumn29.Header.VisiblePosition = 12;
			ultraGridColumn29.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn29});
			ultraGridBand3.Header.Caption = "Semana 2";
			ultraGridBand3.HeaderVisible = true;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance12;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings2});
			this.uGridSemana2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana2.DisplayLayout.Override.ActiveRowAppearance = appearance13;
			this.uGridSemana2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana2.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 7.5F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana2.DisplayLayout.Override.HeaderAppearance = appearance15;
			this.uGridSemana2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance16.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance16.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana2.DisplayLayout.Override.RowAlternateAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana2.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana2.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.uGridSemana2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana2.Location = new System.Drawing.Point(0, 144);
			this.uGridSemana2.Name = "uGridSemana2";
			this.uGridSemana2.Size = new System.Drawing.Size(1024, 104);
			this.uGridSemana2.TabIndex = 731;
			this.uGridSemana2.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana2_AfterRowInsert);
			this.uGridSemana2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana2_KeyDown);
			this.uGridSemana2.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana2_AfterCellUpdate);
			this.uGridSemana2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSemana2_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn25.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn26});
			// 
			// uGridSemana4
			// 
			this.uGridSemana4.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana4.DataSource = this.ultraDataSource6;
			appearance19.BackColor = System.Drawing.Color.White;
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana4.DisplayLayout.Appearance = appearance19;
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 116;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 89;
			ultraGridColumn32.Header.Caption = "Vendedor";
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn32.Width = 190;
			ultraGridColumn33.Header.VisiblePosition = 3;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 46;
			ultraGridColumn34.Header.VisiblePosition = 4;
			ultraGridColumn34.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn34.Width = 75;
			ultraGridColumn35.Header.VisiblePosition = 5;
			ultraGridColumn35.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn35.Width = 75;
			ultraGridColumn36.Header.VisiblePosition = 6;
			ultraGridColumn36.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn36.Width = 75;
			ultraGridColumn37.Header.VisiblePosition = 7;
			ultraGridColumn37.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn37.Width = 75;
			ultraGridColumn38.Header.VisiblePosition = 8;
			ultraGridColumn38.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn38.Width = 75;
			ultraGridColumn39.Header.VisiblePosition = 9;
			ultraGridColumn39.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn39.Width = 75;
			ultraGridColumn40.Header.VisiblePosition = 10;
			ultraGridColumn40.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn40.Width = 75;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance20;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 11;
			ultraGridColumn42.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn42.Header.Caption = "Sector de Carpa";
			ultraGridColumn42.Header.VisiblePosition = 12;
			ultraGridColumn42.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn42});
			ultraGridBand4.Header.Caption = "Semana 4";
			ultraGridBand4.HeaderVisible = true;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance21;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings3});
			this.uGridSemana4.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana4.DisplayLayout.Override.ActiveRowAppearance = appearance22;
			this.uGridSemana4.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana4.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana4.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance23.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana4.DisplayLayout.Override.CardAreaAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance24.FontData.BoldAsString = "True";
			appearance24.FontData.Name = "Arial";
			appearance24.FontData.SizeInPoints = 7.5F;
			appearance24.ForeColor = System.Drawing.Color.White;
			appearance24.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana4.DisplayLayout.Override.HeaderAppearance = appearance24;
			this.uGridSemana4.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance25.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance25.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana4.DisplayLayout.Override.RowAlternateAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana4.DisplayLayout.Override.RowSelectorAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana4.DisplayLayout.Override.SelectedRowAppearance = appearance27;
			this.uGridSemana4.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana4.Location = new System.Drawing.Point(0, 352);
			this.uGridSemana4.Name = "uGridSemana4";
			this.uGridSemana4.Size = new System.Drawing.Size(1024, 104);
			this.uGridSemana4.TabIndex = 733;
			this.uGridSemana4.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana4_AfterRowInsert);
			this.uGridSemana4.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana4_KeyDown);
			this.uGridSemana4.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana4_AfterCellUpdate);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn38.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn39});
			// 
			// uGridSemana3
			// 
			this.uGridSemana3.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana3.DataSource = this.ultraDataSource5;
			appearance28.BackColor = System.Drawing.Color.White;
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana3.DisplayLayout.Appearance = appearance28;
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 116;
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 89;
			ultraGridColumn45.Header.Caption = "Vendedor";
			ultraGridColumn45.Header.VisiblePosition = 2;
			ultraGridColumn45.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn45.Width = 190;
			ultraGridColumn46.Header.VisiblePosition = 3;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 46;
			ultraGridColumn47.Header.VisiblePosition = 4;
			ultraGridColumn47.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn47.Width = 75;
			ultraGridColumn48.Header.VisiblePosition = 5;
			ultraGridColumn48.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn48.Width = 75;
			ultraGridColumn49.Header.VisiblePosition = 6;
			ultraGridColumn49.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn49.Width = 75;
			ultraGridColumn50.Header.VisiblePosition = 7;
			ultraGridColumn50.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn50.Width = 75;
			ultraGridColumn51.Header.VisiblePosition = 8;
			ultraGridColumn51.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn51.Width = 75;
			ultraGridColumn52.Header.VisiblePosition = 9;
			ultraGridColumn52.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn52.Width = 75;
			ultraGridColumn53.Header.VisiblePosition = 10;
			ultraGridColumn53.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn53.Width = 75;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance29;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.VisiblePosition = 11;
			ultraGridColumn55.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn55.Header.Caption = "Sector de Carpa";
			ultraGridColumn55.Header.VisiblePosition = 12;
			ultraGridColumn55.Width = 200;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn55});
			ultraGridBand5.Header.Caption = "Semana 3";
			ultraGridBand5.HeaderVisible = true;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance30;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4});
			this.uGridSemana3.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance31.ForeColor = System.Drawing.Color.Black;
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana3.DisplayLayout.Override.ActiveRowAppearance = appearance31;
			this.uGridSemana3.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance32.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana3.DisplayLayout.Override.CardAreaAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance33.FontData.BoldAsString = "True";
			appearance33.FontData.Name = "Arial";
			appearance33.FontData.SizeInPoints = 7.5F;
			appearance33.ForeColor = System.Drawing.Color.White;
			appearance33.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana3.DisplayLayout.Override.HeaderAppearance = appearance33;
			this.uGridSemana3.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance34.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance34.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana3.DisplayLayout.Override.RowAlternateAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana3.DisplayLayout.Override.RowSelectorAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana3.DisplayLayout.Override.SelectedRowAppearance = appearance36;
			this.uGridSemana3.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana3.Location = new System.Drawing.Point(0, 248);
			this.uGridSemana3.Name = "uGridSemana3";
			this.uGridSemana3.Size = new System.Drawing.Size(1024, 104);
			this.uGridSemana3.TabIndex = 732;
			this.uGridSemana3.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana3_AfterRowInsert);
			this.uGridSemana3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana3_KeyDown);
			this.uGridSemana3.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana3_AfterCellUpdate);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(int);
			ultraDataColumn50.DataType = typeof(int);
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn51.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn52});
			// 
			// uGridSemana5
			// 
			this.uGridSemana5.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana5.DataSource = this.ultraDataSource7;
			appearance37.BackColor = System.Drawing.Color.White;
			appearance37.ForeColor = System.Drawing.Color.Black;
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana5.DisplayLayout.Appearance = appearance37;
			ultraGridColumn56.Header.VisiblePosition = 0;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 116;
			ultraGridColumn57.Header.VisiblePosition = 1;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 89;
			ultraGridColumn58.Header.Caption = "Vendedor";
			ultraGridColumn58.Header.VisiblePosition = 2;
			ultraGridColumn58.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn58.Width = 190;
			ultraGridColumn59.Header.VisiblePosition = 3;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 46;
			ultraGridColumn60.Header.VisiblePosition = 4;
			ultraGridColumn60.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn60.Width = 75;
			ultraGridColumn61.Header.VisiblePosition = 5;
			ultraGridColumn61.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn61.Width = 75;
			ultraGridColumn62.Header.VisiblePosition = 6;
			ultraGridColumn62.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn62.Width = 75;
			ultraGridColumn63.Header.VisiblePosition = 7;
			ultraGridColumn63.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn63.Width = 75;
			ultraGridColumn64.Header.VisiblePosition = 8;
			ultraGridColumn64.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn64.Width = 75;
			ultraGridColumn65.Header.VisiblePosition = 9;
			ultraGridColumn65.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn65.Width = 75;
			ultraGridColumn66.Header.VisiblePosition = 10;
			ultraGridColumn66.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn66.Width = 75;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance38;
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.VisiblePosition = 11;
			ultraGridColumn68.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn68.Header.Caption = "Sector de Carpa";
			ultraGridColumn68.Header.VisiblePosition = 12;
			ultraGridColumn68.Width = 200;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn68});
			ultraGridBand6.Header.Caption = "Semana 5";
			ultraGridBand6.HeaderVisible = true;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance39;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand6.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings5});
			this.uGridSemana5.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance40.ForeColor = System.Drawing.Color.Black;
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana5.DisplayLayout.Override.ActiveRowAppearance = appearance40;
			this.uGridSemana5.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana5.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana5.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance41.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana5.DisplayLayout.Override.CardAreaAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance42.FontData.BoldAsString = "True";
			appearance42.FontData.Name = "Arial";
			appearance42.FontData.SizeInPoints = 7.5F;
			appearance42.ForeColor = System.Drawing.Color.White;
			appearance42.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana5.DisplayLayout.Override.HeaderAppearance = appearance42;
			this.uGridSemana5.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance43.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance43.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana5.DisplayLayout.Override.RowAlternateAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana5.DisplayLayout.Override.RowSelectorAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana5.DisplayLayout.Override.SelectedRowAppearance = appearance45;
			this.uGridSemana5.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana5.Location = new System.Drawing.Point(0, 456);
			this.uGridSemana5.Name = "uGridSemana5";
			this.uGridSemana5.Size = new System.Drawing.Size(1024, 104);
			this.uGridSemana5.TabIndex = 734;
			this.uGridSemana5.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana5_AfterRowInsert);
			this.uGridSemana5.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana5_KeyDown);
			this.uGridSemana5.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana5_AfterCellUpdate);
			// 
			// ultraDataSource7
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
			ultraDataColumn63.DataType = typeof(int);
			ultraDataColumn64.DataType = typeof(System.Decimal);
			ultraDataColumn64.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn65});
			// 
			// dtpMes
			// 
			this.dtpMes.Checked = false;
			this.dtpMes.CustomFormat = "MM/yyyy";
			this.dtpMes.Enabled = false;
			this.dtpMes.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpMes.Location = new System.Drawing.Point(296, 16);
			this.dtpMes.Name = "dtpMes";
			this.dtpMes.ShowUpDown = true;
			this.dtpMes.Size = new System.Drawing.Size(88, 20);
			this.dtpMes.TabIndex = 735;
			this.dtpMes.Value = new System.DateTime(2017, 1, 1, 0, 0, 0, 0);
			this.dtpMes.ValueChanged += new System.EventHandler(this.dtpMes_ValueChanged);
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
			// cmbCom_Ubicacion1
			// 
			this.cmbCom_Ubicacion1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCom_Ubicacion1.DataSource = this.ultraDataSource2;
			ultraGridBand7.AddButtonCaption = "Proyecto";
			ultraGridColumn69.Header.VisiblePosition = 0;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 1;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn69,
																										 ultraGridColumn70});
			this.cmbCom_Ubicacion1.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbCom_Ubicacion1.DisplayMember = "Ubicacion";
			this.cmbCom_Ubicacion1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCom_Ubicacion1.Location = new System.Drawing.Point(408, 136);
			this.cmbCom_Ubicacion1.Name = "cmbCom_Ubicacion1";
			this.cmbCom_Ubicacion1.Size = new System.Drawing.Size(176, 86);
			this.cmbCom_Ubicacion1.TabIndex = 736;
			this.cmbCom_Ubicacion1.ValueMember = "idCom_Ubicacion";
			this.cmbCom_Ubicacion1.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn66.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn66,
																																 ultraDataColumn67});
			// 
			// cmbVendedores
			// 
			this.cmbVendedores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedores.DataSource = this.ultraDataSource3;
			ultraGridColumn71.Header.VisiblePosition = 0;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 57;
			ultraGridColumn72.Header.VisiblePosition = 1;
			ultraGridColumn72.Width = 100;
			ultraGridColumn73.Header.VisiblePosition = 2;
			ultraGridColumn73.Width = 180;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73});
			this.cmbVendedores.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbVendedores.DisplayMember = "Nombre";
			this.cmbVendedores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbVendedores.Location = new System.Drawing.Point(400, 312);
			this.cmbVendedores.Name = "cmbVendedores";
			this.cmbVendedores.Size = new System.Drawing.Size(200, 112);
			this.cmbVendedores.TabIndex = 737;
			this.cmbVendedores.ValueMember = "idPersonal";
			this.cmbVendedores.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn68.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn68,
																																 ultraDataColumn69,
																																 ultraDataColumn70});
			// 
			// cmbCom_Ubicacion2
			// 
			this.cmbCom_Ubicacion2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCom_Ubicacion2.DataSource = this.ultraDataSource2;
			ultraGridBand9.AddButtonCaption = "Proyecto";
			ultraGridColumn74.Header.VisiblePosition = 0;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 1;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn74,
																										 ultraGridColumn75});
			this.cmbCom_Ubicacion2.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbCom_Ubicacion2.DisplayMember = "Ubicacion";
			this.cmbCom_Ubicacion2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCom_Ubicacion2.Location = new System.Drawing.Point(280, 80);
			this.cmbCom_Ubicacion2.Name = "cmbCom_Ubicacion2";
			this.cmbCom_Ubicacion2.Size = new System.Drawing.Size(176, 86);
			this.cmbCom_Ubicacion2.TabIndex = 738;
			this.cmbCom_Ubicacion2.ValueMember = "idCom_Ubicacion";
			this.cmbCom_Ubicacion2.Visible = false;
			// 
			// cmbCom_Ubicacion3
			// 
			this.cmbCom_Ubicacion3.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCom_Ubicacion3.DataSource = this.ultraDataSource2;
			ultraGridBand10.AddButtonCaption = "Proyecto";
			ultraGridColumn76.Header.VisiblePosition = 0;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 1;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn76,
																											ultraGridColumn77});
			this.cmbCom_Ubicacion3.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbCom_Ubicacion3.DisplayMember = "Ubicacion";
			this.cmbCom_Ubicacion3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCom_Ubicacion3.Location = new System.Drawing.Point(648, 56);
			this.cmbCom_Ubicacion3.Name = "cmbCom_Ubicacion3";
			this.cmbCom_Ubicacion3.Size = new System.Drawing.Size(176, 86);
			this.cmbCom_Ubicacion3.TabIndex = 739;
			this.cmbCom_Ubicacion3.ValueMember = "idCom_Ubicacion";
			this.cmbCom_Ubicacion3.Visible = false;
			// 
			// cmbCom_Ubicacion4
			// 
			this.cmbCom_Ubicacion4.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCom_Ubicacion4.DataSource = this.ultraDataSource2;
			ultraGridBand11.AddButtonCaption = "Proyecto";
			ultraGridColumn78.Header.VisiblePosition = 0;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 1;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn78,
																											ultraGridColumn79});
			this.cmbCom_Ubicacion4.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbCom_Ubicacion4.DisplayMember = "Ubicacion";
			this.cmbCom_Ubicacion4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCom_Ubicacion4.Location = new System.Drawing.Point(208, 160);
			this.cmbCom_Ubicacion4.Name = "cmbCom_Ubicacion4";
			this.cmbCom_Ubicacion4.Size = new System.Drawing.Size(176, 86);
			this.cmbCom_Ubicacion4.TabIndex = 740;
			this.cmbCom_Ubicacion4.ValueMember = "idCom_Ubicacion";
			this.cmbCom_Ubicacion4.Visible = false;
			// 
			// cmbCom_Ubicacion5
			// 
			this.cmbCom_Ubicacion5.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCom_Ubicacion5.DataSource = this.ultraDataSource2;
			ultraGridBand12.AddButtonCaption = "Proyecto";
			ultraGridColumn80.Header.VisiblePosition = 0;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 1;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn80,
																											ultraGridColumn81});
			this.cmbCom_Ubicacion5.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbCom_Ubicacion5.DisplayMember = "Ubicacion";
			this.cmbCom_Ubicacion5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCom_Ubicacion5.Location = new System.Drawing.Point(592, 160);
			this.cmbCom_Ubicacion5.Name = "cmbCom_Ubicacion5";
			this.cmbCom_Ubicacion5.Size = new System.Drawing.Size(176, 86);
			this.cmbCom_Ubicacion5.TabIndex = 741;
			this.cmbCom_Ubicacion5.ValueMember = "idCom_Ubicacion";
			this.cmbCom_Ubicacion5.Visible = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(400, 16);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 742;
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
			this.btnConsultar.Location = new System.Drawing.Point(475, 16);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 743;
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
			this.btnEditar.Location = new System.Drawing.Point(556, 15);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 745;
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
			this.btnGuardar.Location = new System.Drawing.Point(635, 15);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 744;
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
			this.btnCancelar.Location = new System.Drawing.Point(792, 15);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 746;
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
			this.btnSalir.Location = new System.Drawing.Point(867, 15);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 747;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtIdHorariosDeCarpa
			// 
			this.txtIdHorariosDeCarpa.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdHorariosDeCarpa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdHorariosDeCarpa.Enabled = false;
			this.txtIdHorariosDeCarpa.Location = new System.Drawing.Point(952, 97);
			this.txtIdHorariosDeCarpa.Name = "txtIdHorariosDeCarpa";
			this.txtIdHorariosDeCarpa.PromptChar = ' ';
			this.txtIdHorariosDeCarpa.Size = new System.Drawing.Size(16, 21);
			this.txtIdHorariosDeCarpa.TabIndex = 748;
			this.txtIdHorariosDeCarpa.Visible = false;
			// 
			// ultraDataSource8
			// 
			ultraDataColumn71.DataType = typeof(int);
			ultraDataColumn72.DataType = typeof(int);
			ultraDataColumn73.DataType = typeof(int);
			ultraDataColumn74.DataType = typeof(int);
			ultraDataColumn75.DataType = typeof(int);
			ultraDataColumn76.DataType = typeof(int);
			ultraDataColumn77.DataType = typeof(int);
			ultraDataColumn78.DataType = typeof(int);
			ultraDataColumn79.DataType = typeof(int);
			ultraDataColumn80.DataType = typeof(int);
			ultraDataColumn81.DataType = typeof(int);
			ultraDataColumn82.DataType = typeof(System.Decimal);
			ultraDataColumn82.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn83});
			// 
			// uGridSemana6
			// 
			this.uGridSemana6.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana6.DataSource = this.ultraDataSource8;
			appearance46.BackColor = System.Drawing.Color.White;
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana6.DisplayLayout.Appearance = appearance46;
			ultraGridColumn82.Header.VisiblePosition = 0;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn82.Width = 116;
			ultraGridColumn83.Header.VisiblePosition = 1;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn83.Width = 89;
			ultraGridColumn84.Header.Caption = "Vendedor";
			ultraGridColumn84.Header.VisiblePosition = 2;
			ultraGridColumn84.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn84.Width = 190;
			ultraGridColumn85.Header.VisiblePosition = 3;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn85.Width = 46;
			ultraGridColumn86.Header.VisiblePosition = 4;
			ultraGridColumn86.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn86.Width = 75;
			ultraGridColumn87.Header.VisiblePosition = 5;
			ultraGridColumn87.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn87.Width = 75;
			ultraGridColumn88.Header.VisiblePosition = 6;
			ultraGridColumn88.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn88.Width = 75;
			ultraGridColumn89.Header.VisiblePosition = 7;
			ultraGridColumn89.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn89.Width = 75;
			ultraGridColumn90.Header.VisiblePosition = 8;
			ultraGridColumn90.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn90.Width = 75;
			ultraGridColumn91.Header.VisiblePosition = 9;
			ultraGridColumn91.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn91.Width = 75;
			ultraGridColumn92.Header.VisiblePosition = 10;
			ultraGridColumn92.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn92.Width = 75;
			ultraGridColumn93.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn93.CellAppearance = appearance47;
			ultraGridColumn93.Format = "#,##0.00";
			ultraGridColumn93.Header.VisiblePosition = 11;
			ultraGridColumn94.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn94.Header.Caption = "Sector de Carpa";
			ultraGridColumn94.Header.VisiblePosition = 12;
			ultraGridColumn94.Width = 200;
			ultraGridBand13.Columns.AddRange(new object[] {
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
																											ultraGridColumn94});
			ultraGridBand13.Header.Caption = "Semana 6";
			ultraGridBand13.HeaderVisible = true;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance48;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand13.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings6});
			this.uGridSemana6.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			appearance49.ForeColor = System.Drawing.Color.Black;
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana6.DisplayLayout.Override.ActiveRowAppearance = appearance49;
			this.uGridSemana6.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana6.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana6.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance50.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana6.DisplayLayout.Override.CardAreaAppearance = appearance50;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance51.FontData.BoldAsString = "True";
			appearance51.FontData.Name = "Arial";
			appearance51.FontData.SizeInPoints = 7.5F;
			appearance51.ForeColor = System.Drawing.Color.White;
			appearance51.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana6.DisplayLayout.Override.HeaderAppearance = appearance51;
			this.uGridSemana6.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance52.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance52.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana6.DisplayLayout.Override.RowAlternateAppearance = appearance52;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana6.DisplayLayout.Override.RowSelectorAppearance = appearance53;
			appearance54.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance54.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana6.DisplayLayout.Override.SelectedRowAppearance = appearance54;
			this.uGridSemana6.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana6.Location = new System.Drawing.Point(0, 560);
			this.uGridSemana6.Name = "uGridSemana6";
			this.uGridSemana6.Size = new System.Drawing.Size(1024, 104);
			this.uGridSemana6.TabIndex = 749;
			this.uGridSemana6.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana6_AfterRowInsert);
			this.uGridSemana6.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana6_KeyDown);
			this.uGridSemana6.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana6_AfterCellUpdate);
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(456, 27);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 750;
			this.txtEstado.Visible = false;
			// 
			// btnAprobar
			// 
			this.btnAprobar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAprobar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnAprobar.CausesValidation = false;
			this.btnAprobar.Enabled = false;
			this.btnAprobar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnAprobar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnAprobar.Image = ((System.Drawing.Image)(resources.GetObject("btnAprobar.Image")));
			this.btnAprobar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAprobar.Location = new System.Drawing.Point(944, 69);
			this.btnAprobar.Name = "btnAprobar";
			this.btnAprobar.TabIndex = 752;
			this.btnAprobar.Text = "Aprobar";
			this.btnAprobar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobar.Visible = false;
			this.btnAprobar.Click += new System.EventHandler(this.btnAprobar_Click);
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
			this.btnAprobarSemana1.Location = new System.Drawing.Point(8, 40);
			this.btnAprobarSemana1.Name = "btnAprobarSemana1";
			this.btnAprobarSemana1.TabIndex = 753;
			this.btnAprobarSemana1.Text = "Aprobar";
			this.btnAprobarSemana1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana1.Click += new System.EventHandler(this.btnAprobarSemana1_Click);
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
			this.btnAprobarSemana2.Location = new System.Drawing.Point(8, 144);
			this.btnAprobarSemana2.Name = "btnAprobarSemana2";
			this.btnAprobarSemana2.TabIndex = 754;
			this.btnAprobarSemana2.Text = "Aprobar";
			this.btnAprobarSemana2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana2.Click += new System.EventHandler(this.btnAprobarSemana2_Click);
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
			this.btnAprobarSemana3.Location = new System.Drawing.Point(8, 248);
			this.btnAprobarSemana3.Name = "btnAprobarSemana3";
			this.btnAprobarSemana3.TabIndex = 755;
			this.btnAprobarSemana3.Text = "Aprobar";
			this.btnAprobarSemana3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana3.Click += new System.EventHandler(this.btnAprobarSemana3_Click);
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
			this.btnAprobarSemana4.Location = new System.Drawing.Point(8, 352);
			this.btnAprobarSemana4.Name = "btnAprobarSemana4";
			this.btnAprobarSemana4.TabIndex = 756;
			this.btnAprobarSemana4.Text = "Aprobar";
			this.btnAprobarSemana4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana4.Click += new System.EventHandler(this.btnAprobarSemana4_Click);
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
			this.btnAprobarSemana5.Location = new System.Drawing.Point(8, 456);
			this.btnAprobarSemana5.Name = "btnAprobarSemana5";
			this.btnAprobarSemana5.TabIndex = 757;
			this.btnAprobarSemana5.Text = "Aprobar";
			this.btnAprobarSemana5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana5.Click += new System.EventHandler(this.btnAprobarSemana5_Click);
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
			this.btnAprobarSemana6.Location = new System.Drawing.Point(8, 560);
			this.btnAprobarSemana6.Name = "btnAprobarSemana6";
			this.btnAprobarSemana6.TabIndex = 758;
			this.btnAprobarSemana6.Text = "Aprobar";
			this.btnAprobarSemana6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana6.Click += new System.EventHandler(this.btnAprobarSemana6_Click);
			// 
			// btnHorarioCarpa
			// 
			this.btnHorarioCarpa.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnHorarioCarpa.CausesValidation = false;
			this.btnHorarioCarpa.Enabled = false;
			this.btnHorarioCarpa.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnHorarioCarpa.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnHorarioCarpa.Image = ((System.Drawing.Image)(resources.GetObject("btnHorarioCarpa.Image")));
			this.btnHorarioCarpa.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnHorarioCarpa.Location = new System.Drawing.Point(714, 15);
			this.btnHorarioCarpa.Name = "btnHorarioCarpa";
			this.btnHorarioCarpa.TabIndex = 759;
			this.btnHorarioCarpa.Text = "Horario";
			this.btnHorarioCarpa.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnHorarioCarpa.Click += new System.EventHandler(this.btnHorarioCarpa_Click);
			// 
			// frmHorariosDeCarpa
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(1006, 594);
			this.Controls.Add(this.btnHorarioCarpa);
			this.Controls.Add(this.btnAprobarSemana6);
			this.Controls.Add(this.btnAprobarSemana5);
			this.Controls.Add(this.btnAprobarSemana4);
			this.Controls.Add(this.btnAprobarSemana3);
			this.Controls.Add(this.btnAprobarSemana2);
			this.Controls.Add(this.btnAprobarSemana1);
			this.Controls.Add(this.btnAprobar);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdHorariosDeCarpa);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.cmbCom_Ubicacion5);
			this.Controls.Add(this.cmbCom_Ubicacion4);
			this.Controls.Add(this.cmbCom_Ubicacion3);
			this.Controls.Add(this.cmbCom_Ubicacion2);
			this.Controls.Add(this.dtpMes);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.uGridSemana5);
			this.Controls.Add(this.uGridSemana4);
			this.Controls.Add(this.uGridSemana3);
			this.Controls.Add(this.uGridSemana2);
			this.Controls.Add(this.uGridSemana1);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbVendedores);
			this.Controls.Add(this.cmbCom_Ubicacion1);
			this.Controls.Add(this.uGridSemana6);
			this.KeyPreview = true;
			this.Name = "frmHorariosDeCarpa";
			this.Text = "Horarios De Carpa";
			this.Load += new System.EventHandler(this.frmHorariosDeCarpa_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCom_Ubicacion5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdHorariosDeCarpa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int idBloqueaTransacciones = 0;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void UnloadMe()
		{
			this.Close();
		}

		bool bNuevo = false;
		bool bEdicion = false;
		int iBodegaPredef = 0;
		private Acceso miAcceso;

		private void frmHorariosDeCarpa_Load(object sender, System.EventArgs e)
		{	
			miAcceso = new Acceso(cdsSeteoF, "0719");
			
			if (!Funcion.Permiso("649", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Planificación de Carpa", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{			
				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
				iBodegaPredef = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario));
			
				this.cmbCom_Ubicacion1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCom_Ubicacion, Ubicacion From Com_Ubicacion");
				this.cmbCom_Ubicacion2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCom_Ubicacion, Ubicacion From Com_Ubicacion");
				this.cmbCom_Ubicacion3.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCom_Ubicacion, Ubicacion From Com_Ubicacion");
				this.cmbCom_Ubicacion4.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCom_Ubicacion, Ubicacion From Com_Ubicacion");
				this.cmbCom_Ubicacion5.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCom_Ubicacion, Ubicacion From Com_Ubicacion");

				FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana1);
				FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana2);
				FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana3);
				FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana4);
				FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana5);
				FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana6);
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
			
			if (this.cmbBodega.ActiveRow != null)
				this.cmbVendedores.DataSource = FuncionesProcedimientos.Lista_VendedoresPlanifCarpa(dtFin, (int)this.cmbBodega.Value, cdsSeteoF);
		
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

		private void OcultaCeldasSemana1()
		{
			DateTime dtInicio = Convert.ToDateTime("01/" + Convert.ToDateTime(dtpMes.Value).Month.ToString() + "/" + Convert.ToDateTime(dtpMes.Value).Year.ToString());
			int iNumDiaSemana = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaNumDiaSemana '{0}'", dtInicio.ToString("yyyyMMdd"))) - 1; 
			if (iNumDiaSemana == 0) iNumDiaSemana = 7;
			iNumDiaSemana = iNumDiaSemana + 4;

			int iContador = 4;

			string sNombreCol = "";

			for (int j = iContador; j <= 10; j++)
			{
				sNombreCol  = this.uGridSemana1.DisplayLayout.Bands[0].Columns[j].Header.Caption;

				if (j < iNumDiaSemana - 1) this.uGridSemana1.DisplayLayout.Bands[0].Columns[sNombreCol].Hidden = true;
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

			iNumDiaSemana = iNumDiaSemana + 4;

			int iContador = 4;		

			if (iSemanas == 5) 
			{
				for (int j = iContador; j <= 10; j++)
				{
					if ((iNumDiaSemana - 1) < j) 
					{
						this.uGridSemana5.DisplayLayout.Bands[0].Columns[j].Hidden = true;						
					}
				}
			}

			if (iSemanas == 6) 
			{
				for (int j = iContador; j <= 10; j++)
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
			iNumDiaSemana = iNumDiaSemana + 4;
			
			int iDia = 1;		
			int iContador = 4;	
			int iSemanas = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CalculaSemanasXMes '{0}'", dtInicio.ToString("yyyyMMdd")));

			for (int j = iContador; j <= 10; j++)
			{				
				if (!this.uGridSemana1.DisplayLayout.Bands[0].Columns[j].Hidden)
				{
					this.uGridSemana1.DisplayLayout.Bands[0].Columns[j].Header.Caption = this.uGridSemana1.DisplayLayout.Bands[0].Columns[j].Header.Caption + "-" + iDia.ToString();
					iDia++;
				}
			}

			iContador = 4;
			for (int j = iContador; j <= 10; j++)
			{				
				if (!this.uGridSemana2.DisplayLayout.Bands[0].Columns[j].Hidden)
				{
					this.uGridSemana2.DisplayLayout.Bands[0].Columns[j].Header.Caption = this.uGridSemana2.DisplayLayout.Bands[0].Columns[j].Header.Caption + "-" + iDia.ToString();
					iDia++;
				}
			}

			iContador = 4;
			for (int j = iContador; j <= 10; j++)
			{	
				if (!this.uGridSemana3.DisplayLayout.Bands[0].Columns[j].Hidden)
				{
					this.uGridSemana3.DisplayLayout.Bands[0].Columns[j].Header.Caption = this.uGridSemana3.DisplayLayout.Bands[0].Columns[j].Header.Caption + "-" + iDia.ToString();
					iDia++;
				}
			}

			iContador = 4;
			for (int j = iContador; j <= 10; j++)
			{
				if (!this.uGridSemana4.DisplayLayout.Bands[0].Columns[j].Hidden)
				{
					if (iDia <= UDia)
						this.uGridSemana4.DisplayLayout.Bands[0].Columns[j].Header.Caption = this.uGridSemana4.DisplayLayout.Bands[0].Columns[j].Header.Caption + "-" + iDia.ToString();
					iDia++;
				}
			}
			
			iContador = 4;
			for (int j = iContador; j <= 10; j++)
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

				iContador = 4;
				for (int j = iContador; j <= 10; j++)
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

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{	
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 5", idBloqueaTransacciones, (int)this.txtIdHorariosDeCarpa.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			this.cmbBodega.Enabled = false;
			this.dtpMes.Enabled = false;
			this.dtpMes.Value = DateTime.Parse("01/01/2017");

			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana1, 0, 1, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana2, 0, 2, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana3, 0, 3, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana4, 0, 4, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana5, 0, 5, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana6, 0, 6, 1, cdsSeteoF);

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

			this.txtIdHorariosDeCarpa.Value = 0;
			this.cmbBodega.Value = System.DBNull.Value;
			this.txtEstado.Value = 0;
			idBloqueaTransacciones = 0;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAprobar.Enabled = false;
			this.btnCancelar.Enabled = false;

			this.btnAprobarSemana1.Enabled = false;
			this.btnAprobarSemana2.Enabled = false;
			this.btnAprobarSemana3.Enabled = false;
			this.btnAprobarSemana4.Enabled = false;
			this.btnAprobarSemana5.Enabled = false;
			this.btnAprobarSemana6.Enabled = false;

			this.btnHorarioCarpa.Enabled = false;
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
			this.btnAprobar.Enabled = false;
			this.btnCancelar.Enabled = true;

			this.btnAprobarSemana1.Enabled = false;
			this.btnAprobarSemana2.Enabled = false;
			this.btnAprobarSemana3.Enabled = false;
			this.btnAprobarSemana4.Enabled = false;
			this.btnAprobarSemana5.Enabled = false;
			this.btnAprobarSemana6.Enabled = false;

			if (iBodegaPredef == 0) 
			{
				this.cmbBodega.Enabled = true;
				this.cmbBodega.Value = 12;
				this.cmbBodega.Focus();
			}
			else
			{
				this.cmbBodega.Value = iBodegaPredef;				
				this.dtpMes.Focus();
			}

			this.dtpMes.Value = DateTime.Today;//DateTime.Parse("30/04/2017");//

			bNuevo = true;
			bEdicion = true;
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

		private void GuardaEstadoHorariosDeCarpaPorDias (int iSemana, int iDia)
		{
			//			string sSQL = string.Format("Exec GuardaEstadoHorariosDeCarpaPorDias {0}, {1}, '{2}', '{3}', {4}", (int)this.txtIdHorariosDeCarpa.Value, iSemana, iDia, dtDesde, dtHasta);
			//			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			string sSQLVal = string.Format("Exec ValidaHorariosDeCarpa {0}, '{1}', 1", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtpMes.Value).ToString("yyyyMMdd"));
			if (Funcion.bEscalarSQL(cdsSeteoF, sSQLVal, true) && bNuevo)
			{
				MessageBox.Show("Ya existe una planificacion de carpa del local y periodo seleccionado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);        				
				this.dtpMes.Focus();
				return;
			}

			string sSQL = string.Format("Exec GuardaHorariosCarpa {0}, {1}, '{2}', '', 1", 
				(int)this.txtIdHorariosDeCarpa.Value, (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtpMes.Value).ToString("yyyyMMdd"));
			this.txtIdHorariosDeCarpa.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (this.uGridSemana1.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana1, (int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF);			
			if (this.uGridSemana2.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana2, (int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF);
			if (this.uGridSemana3.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana3, (int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF);
			if (this.uGridSemana4.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana4, (int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF);
			if (this.uGridSemana5.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana5, (int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF);
			if (this.uGridSemana6.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana6, (int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF);

			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana1, (int)this.txtIdHorariosDeCarpa.Value, 1, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana2, (int)this.txtIdHorariosDeCarpa.Value, 2, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana3, (int)this.txtIdHorariosDeCarpa.Value, 3, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana4, (int)this.txtIdHorariosDeCarpa.Value, 4, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana5, (int)this.txtIdHorariosDeCarpa.Value, 5, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana6, (int)this.txtIdHorariosDeCarpa.Value, 6, 1, cdsSeteoF);

			if (idBloqueaTransacciones > 0)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 5", idBloqueaTransacciones, (int)this.txtIdHorariosDeCarpa.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			this.cmbBodega.Enabled = false;
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
			{
				this.btnEditar.Enabled = true;
				this.btnAprobar.Enabled = true;
			}
			
			this.btnAprobarSemana1.Enabled = true;
			this.btnAprobarSemana2.Enabled = true;
			this.btnAprobarSemana3.Enabled = true;
			this.btnAprobarSemana4.Enabled = true;
			this.btnAprobarSemana5.Enabled = true;
			this.btnAprobarSemana6.Enabled = true;			
			
			this.btnHorarioCarpa.Enabled = false;

			this.btnCancelar.Enabled = true;
		}

		private void Consultar(int idHorarioCarpa)
		{			
			DateTime dtMes = DateTime.Today;

			string sSQL = string.Format("Exec ConsultaHorariosCarpaIndivisual {0}, 1", idHorarioCarpa);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdHorariosDeCarpa.Value = idHorarioCarpa;
				this.cmbBodega.Value = dr.GetInt32(1);
				dtMes = dr.GetDateTime(2);
				this.txtEstado.Value = dr.GetInt32(4);
				//this.lblUsuario.Text = dr.GetString(5);
			}
			dr.Close();
		
			this.dtpMes.Value = dtMes; 

			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana1, idHorarioCarpa, 1, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana2, idHorarioCarpa, 2, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana3, idHorarioCarpa, 3, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana4, idHorarioCarpa, 4, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana5, idHorarioCarpa, 5, 1, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana6, idHorarioCarpa, 6, 1, cdsSeteoF);

			this.uGridSemana1.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana1.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF, 1);
			this.uGridSemana2.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana2.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF, 2);
			this.uGridSemana3.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana3.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF, 3);
			this.uGridSemana4.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana4.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF, 4);
			this.uGridSemana5.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana5.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF, 5);
			this.uGridSemana6.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana6.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdHorariosDeCarpa.Value, 1, cdsSeteoF, 6);

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

			using (frmBuscaHorariosCarpa Busqueda = new frmBuscaHorariosCarpa(iBodegaPredef))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consultar((int)Busqueda.uGridHorarios.ActiveRow.Cells["idHorariosDeCarpa"].Value);
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAprobar.Enabled = false;
			this.btnCancelar.Enabled = true;
			
			this.btnAprobarSemana1.Enabled = false;
			this.btnAprobarSemana2.Enabled = false;
			this.btnAprobarSemana3.Enabled = false;
			this.btnAprobarSemana4.Enabled = false;
			this.btnAprobarSemana5.Enabled = false;
			this.btnAprobarSemana6.Enabled = false;

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 1", 1, (int)this.txtIdHorariosDeCarpa.Value)) > 0)
			{
				if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana1);	
			}
			else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana1);
			
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 1", 2, (int)this.txtIdHorariosDeCarpa.Value)) > 0)
			{
				if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana2);	
			}
			else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana2);

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 1", 3, (int)this.txtIdHorariosDeCarpa.Value)) > 0)
			{
				if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana3);
			}
			else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana3);

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 1", 4, (int)this.txtIdHorariosDeCarpa.Value)) > 0)
			{
				if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana4);
			}
			else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana4);

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 1", 5, (int)this.txtIdHorariosDeCarpa.Value)) > 0)
			{
				if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana5);	
			}
			else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana5);

			if (this.uGridSemana6.Rows.Count > 0) 
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 1", 6, (int)this.txtIdHorariosDeCarpa.Value)) > 0)
				{
					if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana6);	
				}
				else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana6);
			}

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 5", (int)this.txtIdHorariosDeCarpa.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			this.btnHorarioCarpa.Enabled = true;

			bNuevo = false;
			bEdicion = true;
		}

		private void uGridSemana1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana1, 1, 1);
		}

		private void uGridSemana2_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana2, 2, 1);
		}

		private void uGridSemana3_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana3, 3, 1);
		}

		private void uGridSemana4_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana4, 4, 1);
		}

		private void uGridSemana5_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana5, 5, 1);
		}

		private void uGridSemana6_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridSemana6);
		}

		private void uGridSemana6_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana6, 6, 1);
		}
		
		private void uGridSemana1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnAprobar_Click(object sender, System.EventArgs e)
		{
			int iEstado = 1;

			string sSQL = string.Format("Exec ActualizaEstadoHorariosCarpa {0}, {1}", iEstado, (int)this.txtIdHorariosDeCarpa.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
	
			this.txtEstado.Value = 1;
			this.btnEditar.Enabled = false;
			this.btnAprobar.Enabled = false;
		}
		
		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) this.dtpMes_ValueChanged(sender, e);
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
				string.Format("Exec RetornaIdASHC {0}, {1}, 1", iSemana, (int)this.txtIdHorariosDeCarpa.Value));

			if (idAprobacionSemanalHorariosDeCarpa == 0 && !miAcceso.BAprobarCarpa)
			{
				MessageBox.Show(string.Format("La semana {0} de carpa no ha sido Aprobada\n\nConsulte con el Jefe Nacional De Ventas", iSemana), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 5", (int)this.txtIdHorariosDeCarpa.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			using (frmNotasAprobacionCarpa Notas = new frmNotasAprobacionCarpa((int)this.txtIdHorariosDeCarpa.Value, iSemana, miAcceso.BAprobarCarpa, 1))
			{
				if (DialogResult.OK == Notas.ShowDialog())
				{
					try
					{							
						string sSQL = string.Format("Exec GuardaAprobacionSemanalHorariosDeCarpa {0}, {1}, {2}, '{3}', 1", 
							idAprobacionSemanalHorariosDeCarpa, (int)this.txtIdHorariosDeCarpa.Value, iSemana, Notas.txtNotas.Text.ToString());
						idAprobacionSemanalHorariosDeCarpa = Funcion.iEscalarSQL(cdsSeteoF, sSQL);						
					}
					catch(Exception Exc)
					{
						MessageBox.Show(string.Format("Error al Aprobar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							
					}
				}
			}

			sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 5", idBloqueaTransacciones, (int)this.txtIdHorariosDeCarpa.Value);
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

		private void uGridSemana1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Lunes" || e.Cell.Column.ToString() == "Martes" || e.Cell.Column.ToString() == "Miercoles" || e.Cell.Column.ToString() == "Jueves" || 
				e.Cell.Column.ToString() == "Viernes" || e.Cell.Column.ToString() == "Sabado" || e.Cell.Column.ToString() == "Domingo")
				CalculaPresupuesto(sender, e, this.uGridSemana1);
		}

		private void CalculaPresupuesto(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e, Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			decimal dPresupuesto = 1000.00m;
			decimal dPresupuestoTotal = 0.00m;
			e.Cell.Row.Cells["Presupuesto"].Value = 0.00m;
			string nombreColumna = "";
			
			for (int j = 0; j <= 10; j++)
			{
				if (!uGrid.DisplayLayout.Bands[0].Columns[j].Hidden)
				{
					nombreColumna = uGrid.DisplayLayout.Bands[0].Columns[j].Key.ToString();//.Header.Caption;
					
					if (nombreColumna == "Lunes" || nombreColumna == "Martes" || nombreColumna == "Miercoles" || nombreColumna == "Jueves" || nombreColumna == "Viernes" || nombreColumna == "Sabado" || nombreColumna == "Domingo")
					{
						if (e.Cell.Row.Cells[nombreColumna].Value != System.DBNull.Value)			
							if (int.Parse(e.Cell.Row.Cells[nombreColumna].Value.ToString()) == 2)
								e.Cell.Row.Cells["Presupuesto"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Presupuesto"].Value) + dPresupuesto;
					}
				}
			}

			//e.Cell.Row.Cells["Presupuesto"].Value = dPresupuestoTotal;
		}

		private void btnHorarioCarpa_Click(object sender, System.EventArgs e)
		{
			using (frmFechaHoraHorariosCarpa ESCarpa = new frmFechaHoraHorariosCarpa(DateTime.Today, (int)this.txtIdHorariosDeCarpa.Value))
			{
				if (DialogResult.OK == ESCarpa.ShowDialog())
				{
					try
					{
						DateTime dtEntrada = DateTime.Today;
						bool bHasta = false;

						if (ESCarpa.dtHasta.Value != System.DBNull.Value) 
						{
							dtEntrada = Convert.ToDateTime(ESCarpa.dtHasta.Value);
							bHasta = true;
						}

						string sSQL = string.Format("Exec GuardaEstadoHorariosDeCarpaPorDias {0}, {1}, '{2}', '{3}', '{4}', '{5}', {6}, {7}", 
							(int)this.txtIdHorariosDeCarpa.Value, ESCarpa.iDia, 
							Convert.ToDateTime(ESCarpa.dtDesde.Value).ToString("yyyyMMdd HH:mm"), ESCarpa.txtObservacionesSalida.Text.ToString(), 
							dtEntrada.ToString("yyyyMMdd HH:mm"), ESCarpa.txtObservacionesEntrada.Text.ToString(), 
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

		private void uGridSemana2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridSemana2_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Lunes" || e.Cell.Column.ToString() == "Martes" || e.Cell.Column.ToString() == "Miercoles" || e.Cell.Column.ToString() == "Jueves" || 
				e.Cell.Column.ToString() == "Viernes" || e.Cell.Column.ToString() == "Sabado" || e.Cell.Column.ToString() == "Domingo")
				CalculaPresupuesto(sender, e, this.uGridSemana2);
		}

		private void uGridSemana3_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Lunes" || e.Cell.Column.ToString() == "Martes" || e.Cell.Column.ToString() == "Miercoles" || e.Cell.Column.ToString() == "Jueves" || 
				e.Cell.Column.ToString() == "Viernes" || e.Cell.Column.ToString() == "Sabado" || e.Cell.Column.ToString() == "Domingo")
				CalculaPresupuesto(sender, e, this.uGridSemana3);
		}

		private void uGridSemana4_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Lunes" || e.Cell.Column.ToString() == "Martes" || e.Cell.Column.ToString() == "Miercoles" || e.Cell.Column.ToString() == "Jueves" || 
				e.Cell.Column.ToString() == "Viernes" || e.Cell.Column.ToString() == "Sabado" || e.Cell.Column.ToString() == "Domingo")
				CalculaPresupuesto(sender, e, this.uGridSemana4);
		}

		private void uGridSemana5_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Lunes" || e.Cell.Column.ToString() == "Martes" || e.Cell.Column.ToString() == "Miercoles" || e.Cell.Column.ToString() == "Jueves" || 
				e.Cell.Column.ToString() == "Viernes" || e.Cell.Column.ToString() == "Sabado" || e.Cell.Column.ToString() == "Domingo")
				CalculaPresupuesto(sender, e, this.uGridSemana5);
		}

		private void uGridSemana6_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Lunes" || e.Cell.Column.ToString() == "Martes" || e.Cell.Column.ToString() == "Miercoles" || e.Cell.Column.ToString() == "Jueves" || 
				e.Cell.Column.ToString() == "Viernes" || e.Cell.Column.ToString() == "Sabado" || e.Cell.Column.ToString() == "Domingo")
				CalculaPresupuesto(sender, e, this.uGridSemana6);
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
