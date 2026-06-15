using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteRDEP.
	/// </summary>
	public class frmReporteRDEP : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdRolNomina;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEDepartamento;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label33;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECargo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnio;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECentroCosto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridRoles;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteRDEP()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idModulo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Com_IdCargo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteRDEP));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellidos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Genero");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo_Sectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Actividad_Sectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Enero");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Enero");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Enero");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Enero");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Enero");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Enero");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Febrero");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Febrero");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Febrero");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Febrero");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Febrero");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Febrero");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Marzo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Marzo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Marzo");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Marzo");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Marzo");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Marzo");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Abril");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Abril");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Abril");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Abril");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Abril");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Abril");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Mayo");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Mayo");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Mayo");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Mayo");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Mayo");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Mayo");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Junio");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Junio");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Junio");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Junio");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Junio");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Junio");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Julio");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Julio");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Julio");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Julio");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Julio");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Julio");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Agosto");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Agosto");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Agosto");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Agosto");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Agosto");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Agosto");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Septiembre");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Septiembre");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Septiembre");
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Septiembre");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Septiembre");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Septiembre");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Octubre");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Octubre");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Octubre");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Octubre");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Octubre");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Octubre");
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Noviembre");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Noviembre");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Noviembre");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Noviembre");
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Noviembre");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Noviembre");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Enero");
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Enero");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Febrero");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Febrero");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Marzo");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Marzo");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Abril");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Abril");
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Mayo");
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Mayo");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Junio");
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Junio");
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Julio");
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Julio");
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Agosto");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Agosto");
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Septiembre");
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Septiembre");
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Octubre");
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Octubre");
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Noviembre");
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Noviembre");
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo_Diciembre");
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo_Diciembre");
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base_Periodo_Diciembre");
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo_Periodo_Diciembre");
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sobre_Sueldo_Periodo_Diciembre");
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte_Patronal_Diciembre");
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_Acum_Diciembre");
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_R_No_Acum_Diciembre");
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Apellidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Genero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo_Sectorial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Actividad_Sectorial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Departamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Enero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Enero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Enero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Enero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Enero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Enero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Febrero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Febrero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Febrero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Febrero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Febrero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Febrero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Marzo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Marzo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Marzo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Marzo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Marzo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Marzo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Abril");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Abril");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Abril");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Abril");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Abril");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Abril");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Mayo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Mayo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Mayo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Mayo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Mayo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Mayo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Junio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Junio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Junio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Junio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Junio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Junio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Julio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Julio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Julio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Julio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Julio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Julio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Agosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Agosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Agosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Agosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Agosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Agosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Septiembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Septiembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Septiembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Septiembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Septiembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Septiembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Octubre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Octubre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Octubre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Octubre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Octubre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Octubre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Noviembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Noviembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Noviembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Noviembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Noviembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Noviembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Enero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Enero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Febrero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Febrero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Marzo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Marzo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Abril");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Abril");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Mayo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Mayo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Junio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Junio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Julio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Julio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Agosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn98 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Agosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn99 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Septiembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn100 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Septiembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn101 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Octubre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn102 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Octubre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn103 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Noviembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn104 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Noviembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn105 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo_Diciembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn106 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo_Diciembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn107 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base_Periodo_Diciembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn108 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo_Periodo_Diciembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn109 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sobre_Sueldo_Periodo_Diciembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn110 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte_Patronal_Diciembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn111 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_Acum_Diciembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn112 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("F_R_No_Acum_Diciembre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn113 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Revisado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn114 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleRolNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn115 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn116 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn117 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn118 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn119 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Apellidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn120 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn121 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn122 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Salida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn123 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasTrabajados");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn124 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn125 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn126 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mujer");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn127 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Discapacitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn128 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDePago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn129 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn130 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn131 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn132 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn133 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn134 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn135 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bonos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn136 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BonoAdicional");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn137 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comisiones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn138 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Alimentacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn139 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Movilizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn140 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Spiff");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn141 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VHOrdinaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn142 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NHOrdinarias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn143 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HOrdinarias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn144 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VHExtraOrdinaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn145 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NHExtraOrdinarias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn146 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HExtraOrdinarias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn147 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionReal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn148 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DerechoFondosReserva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn149 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FondosReservaAcumula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn150 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FRAcumula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn151 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FRNoAcumula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn152 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalIngresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn153 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BaseIESS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn154 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecAcumula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn155 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn156 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn157 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vacaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn158 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AportePatronal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn159 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoIESS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn160 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Extension");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn161 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrestamoQuirografario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn162 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrestamoHipotecario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn163 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PagoIESS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn164 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComprasCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn165 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComprasCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn166 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FaltantesCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn167 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Multas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn168 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AnticipoComisiones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn169 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AnticipoSueldos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn170 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SeguroMedico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn171 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Farmacia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn172 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OtrosDescuentos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn173 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Faltas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn174 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetencionJudicial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn175 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalEgresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn176 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn177 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn178 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesFC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn179 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesMultas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn180 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesAnticipos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn181 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesOtros");
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdRolNomina = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbNEDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label13 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label33 = new System.Windows.Forms.Label();
			this.cmbNECargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtAnio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.lblContador = new System.Windows.Forms.Label();
			this.cmbNECentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.uGridRoles = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolNomina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRoles)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(360, 8);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 21);
			this.txtIdAsiento.TabIndex = 790;
			this.txtIdAsiento.Visible = false;
			// 
			// txtIdRolNomina
			// 
			this.txtIdRolNomina.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdRolNomina.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdRolNomina.Enabled = false;
			this.txtIdRolNomina.Location = new System.Drawing.Point(384, 8);
			this.txtIdRolNomina.Name = "txtIdRolNomina";
			this.txtIdRolNomina.PromptChar = ' ';
			this.txtIdRolNomina.Size = new System.Drawing.Size(16, 21);
			this.txtIdRolNomina.TabIndex = 789;
			this.txtIdRolNomina.Visible = false;
			// 
			// cmbNEDepartamento
			// 
			this.cmbNEDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEDepartamento.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 205;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbNEDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbNEDepartamento.DisplayMember = "Descripcion";
			this.cmbNEDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEDepartamento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEDepartamento.Location = new System.Drawing.Point(416, 40);
			this.cmbNEDepartamento.Name = "cmbNEDepartamento";
			this.cmbNEDepartamento.Size = new System.Drawing.Size(208, 22);
			this.cmbNEDepartamento.TabIndex = 788;
			this.cmbNEDepartamento.ValueMember = "idDepartamento";
			this.cmbNEDepartamento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEDepartamento_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(8, 72);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(35, 16);
			this.label13.TabIndex = 787;
			this.label13.Text = "Cargo";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(328, 40);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(76, 16);
			this.label12.TabIndex = 782;
			this.label12.Text = "Departamento";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(328, 72);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(84, 16);
			this.label7.TabIndex = 779;
			this.label7.Text = "Cedula/Nombre";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(8, 40);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(88, 16);
			this.label33.TabIndex = 780;
			this.label33.Text = "Centro De Costo";
			// 
			// cmbNECargo
			// 
			this.cmbNECargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECargo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridBand2.AddButtonCaption = "RolTipo";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 208;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbNECargo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNECargo.DisplayMember = "Descripcion";
			this.cmbNECargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECargo.Location = new System.Drawing.Point(104, 72);
			this.cmbNECargo.Name = "cmbNECargo";
			this.cmbNECargo.Size = new System.Drawing.Size(208, 21);
			this.cmbNECargo.TabIndex = 786;
			this.cmbNECargo.ValueMember = "idCargo";
			this.cmbNECargo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNECargo_KeyDown);
			// 
			// txtAnio
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAnio.Appearance = appearance1;
			this.txtAnio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAnio.FormatString = "";
			this.txtAnio.Location = new System.Drawing.Point(104, 8);
			this.txtAnio.MaskInput = "nnnn";
			this.txtAnio.Name = "txtAnio";
			this.txtAnio.PromptChar = ' ';
			this.txtAnio.Size = new System.Drawing.Size(96, 21);
			this.txtAnio.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAnio.TabIndex = 785;
			this.txtAnio.ValueChanged += new System.EventHandler(this.txtAnio_ValueChanged);
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel4.Location = new System.Drawing.Point(8, 16);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(23, 15);
			this.ultraLabel4.TabIndex = 784;
			this.ultraLabel4.Text = "Año";
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(272, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 18);
			this.lblContador.TabIndex = 783;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbNECentroCosto
			// 
			this.cmbNECentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridBand3.AddButtonCaption = "RolTipo";
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 85;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Width = 208;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbNECentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbNECentroCosto.DisplayMember = "Nombre";
			this.cmbNECentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECentroCosto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECentroCosto.Location = new System.Drawing.Point(104, 40);
			this.cmbNECentroCosto.Name = "cmbNECentroCosto";
			this.cmbNECentroCosto.Size = new System.Drawing.Size(208, 22);
			this.cmbNECentroCosto.TabIndex = 781;
			this.cmbNECentroCosto.ValueMember = "idProyecto";
			this.cmbNECentroCosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNECentroCosto_KeyDown);
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(416, 72);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(208, 19);
			this.txtBuscar.TabIndex = 778;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Enabled = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(640, 72);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 777;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// uGridRoles
			// 
			this.uGridRoles.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridRoles.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridRoles.DataSource = this.ultraDataSource3;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridRoles.DisplayLayout.Appearance = appearance2;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance3;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 200;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 200;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance4;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 55;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Width = 70;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "Codigo Sectorial";
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.Width = 115;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Actividad Sectorial";
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Width = 150;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "Centro De Costo";
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridColumn16.Width = 200;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Width = 150;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 12;
			ultraGridColumn18.Width = 150;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.Caption = "Tipo Pago";
			ultraGridColumn19.Header.VisiblePosition = 8;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 9;
			ultraGridColumn20.Width = 120;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance5;
			ultraGridColumn21.Header.VisiblePosition = 10;
			ultraGridColumn21.Width = 80;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance6;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "Base Periodo Enero";
			ultraGridColumn22.Header.VisiblePosition = 13;
			ultraGridColumn22.Width = 180;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance7;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "Sueldo Periodo Enero";
			ultraGridColumn23.Header.VisiblePosition = 14;
			ultraGridColumn23.Width = 180;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance8;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "Sobre Sueldo Periodo Enero";
			ultraGridColumn24.Header.VisiblePosition = 15;
			ultraGridColumn24.Width = 180;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance9;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.Caption = "Aporte Personal Enero";
			ultraGridColumn25.Header.VisiblePosition = 16;
			ultraGridColumn25.Width = 180;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance10;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "FR Acum Enero";
			ultraGridColumn26.Header.VisiblePosition = 17;
			ultraGridColumn26.Width = 180;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance11;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.Caption = "FR No Acum Enero";
			ultraGridColumn27.Header.VisiblePosition = 18;
			ultraGridColumn27.Width = 180;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance12;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.Caption = "Base Periodo Febrero";
			ultraGridColumn28.Header.VisiblePosition = 21;
			ultraGridColumn28.Width = 180;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance13;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.Caption = "Sueldo Periodo Febrero";
			ultraGridColumn29.Header.VisiblePosition = 22;
			ultraGridColumn29.Width = 180;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance14;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "Sobre Sueldo Periodo Febrero";
			ultraGridColumn30.Header.VisiblePosition = 23;
			ultraGridColumn30.Width = 180;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance15;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.Caption = "Aporte Personal Febrero";
			ultraGridColumn31.Header.VisiblePosition = 24;
			ultraGridColumn31.Width = 180;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance16;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.Caption = "FR Acum Febrero";
			ultraGridColumn32.Header.VisiblePosition = 25;
			ultraGridColumn32.Width = 180;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance17;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.Caption = "FR No Acum Febrero";
			ultraGridColumn33.Header.VisiblePosition = 26;
			ultraGridColumn33.Width = 180;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance18;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.Caption = "Base Periodo Marzo";
			ultraGridColumn34.Header.VisiblePosition = 29;
			ultraGridColumn34.Width = 180;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance19;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.Caption = "Sueldo Periodo Marzo";
			ultraGridColumn35.Header.VisiblePosition = 30;
			ultraGridColumn35.Width = 180;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance20;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.Caption = "Sobre Sueldo Periodo Marzo";
			ultraGridColumn36.Header.VisiblePosition = 31;
			ultraGridColumn36.Width = 180;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance21;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.Caption = "Aporte Personal Marzo";
			ultraGridColumn37.Header.VisiblePosition = 32;
			ultraGridColumn37.Width = 180;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance22;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "FR Acum Marzo";
			ultraGridColumn38.Header.VisiblePosition = 33;
			ultraGridColumn38.Width = 180;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance23;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.Caption = "FR No Acum Marzo";
			ultraGridColumn39.Header.VisiblePosition = 34;
			ultraGridColumn39.Width = 180;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance24;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.Caption = "Base Periodo Abril";
			ultraGridColumn40.Header.VisiblePosition = 37;
			ultraGridColumn40.Width = 180;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance25;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.Caption = "Sueldo Periodo Abril";
			ultraGridColumn41.Header.VisiblePosition = 38;
			ultraGridColumn41.Width = 180;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance26;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.Caption = "Sobre Sueldo Periodo Abril";
			ultraGridColumn42.Header.VisiblePosition = 39;
			ultraGridColumn42.Width = 180;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance27;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.Caption = "Aporte Personal Abril";
			ultraGridColumn43.Header.VisiblePosition = 40;
			ultraGridColumn43.Width = 180;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance28;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.Caption = "FR Acum Abril";
			ultraGridColumn44.Header.VisiblePosition = 41;
			ultraGridColumn44.Width = 180;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance29;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.Caption = "FR No Acum Abril";
			ultraGridColumn45.Header.VisiblePosition = 42;
			ultraGridColumn45.Width = 180;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance30;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.Caption = "Base Periodo Mayo";
			ultraGridColumn46.Header.VisiblePosition = 45;
			ultraGridColumn46.Width = 180;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance31;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.Caption = "Sueldo Periodo Mayo";
			ultraGridColumn47.Header.VisiblePosition = 46;
			ultraGridColumn47.Width = 180;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance32;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.Caption = "Sobre Sueldo Periodo Mayo";
			ultraGridColumn48.Header.VisiblePosition = 47;
			ultraGridColumn48.Width = 180;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance33;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.Caption = "Aporte Personal Mayo";
			ultraGridColumn49.Header.VisiblePosition = 48;
			ultraGridColumn49.Width = 180;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance34;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.Caption = "FR Acum Mayo";
			ultraGridColumn50.Header.VisiblePosition = 49;
			ultraGridColumn50.Width = 180;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance35;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.Caption = "FR No Acum Mayo";
			ultraGridColumn51.Header.VisiblePosition = 50;
			ultraGridColumn51.Width = 180;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance36;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.Caption = "Base Periodo Junio";
			ultraGridColumn52.Header.VisiblePosition = 53;
			ultraGridColumn52.Width = 180;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance37;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.Caption = "Sueldo Periodo Junio";
			ultraGridColumn53.Header.VisiblePosition = 54;
			ultraGridColumn53.Width = 180;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance38;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.Caption = "Sobre Sueldo Periodo Junio";
			ultraGridColumn54.Header.VisiblePosition = 55;
			ultraGridColumn54.Width = 180;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance39;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.Caption = "Aporte Personal Junio";
			ultraGridColumn55.Header.VisiblePosition = 56;
			ultraGridColumn55.Width = 180;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance40;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.Caption = "FR Acum Junio";
			ultraGridColumn56.Header.VisiblePosition = 57;
			ultraGridColumn56.Width = 180;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn57.CellAppearance = appearance41;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.Caption = "FR No Acum Junio";
			ultraGridColumn57.Header.VisiblePosition = 58;
			ultraGridColumn57.Width = 180;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance42;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.Caption = "Base Periodo Julio";
			ultraGridColumn58.Header.VisiblePosition = 61;
			ultraGridColumn58.Width = 180;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance43;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.Caption = "Sueldo Periodo Julio";
			ultraGridColumn59.Header.VisiblePosition = 62;
			ultraGridColumn59.Width = 180;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance44;
			ultraGridColumn60.Format = "#,##0.00";
			ultraGridColumn60.Header.Caption = "Sobre Sueldo Periodo Julio";
			ultraGridColumn60.Header.VisiblePosition = 63;
			ultraGridColumn60.Width = 180;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance45;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.Caption = "Aporte Personal Julio";
			ultraGridColumn61.Header.VisiblePosition = 64;
			ultraGridColumn61.Width = 180;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance46;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.Caption = "FR Acum Julio";
			ultraGridColumn62.Header.VisiblePosition = 65;
			ultraGridColumn62.Width = 180;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance47;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.Caption = "FR No Acum Julio";
			ultraGridColumn63.Header.VisiblePosition = 66;
			ultraGridColumn63.Width = 180;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance48;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.Caption = "Base Periodo Agosto";
			ultraGridColumn64.Header.VisiblePosition = 69;
			ultraGridColumn64.Width = 180;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance49;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.Caption = "Sueldo Periodo Agosto";
			ultraGridColumn65.Header.VisiblePosition = 70;
			ultraGridColumn65.Width = 180;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance50;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.Caption = "Sobre Sueldo Periodo Agosto";
			ultraGridColumn66.Header.VisiblePosition = 71;
			ultraGridColumn66.Width = 180;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance51;
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.Caption = "Aporte Personal Agosto";
			ultraGridColumn67.Header.VisiblePosition = 72;
			ultraGridColumn67.Width = 180;
			ultraGridColumn68.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance52;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.Caption = "FR Acum Agosto";
			ultraGridColumn68.Header.VisiblePosition = 73;
			ultraGridColumn68.Width = 180;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance53;
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.Caption = "FR No Acum Agosto";
			ultraGridColumn69.Header.VisiblePosition = 74;
			ultraGridColumn69.Width = 180;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance54;
			ultraGridColumn70.Format = "#,##0.00";
			ultraGridColumn70.Header.Caption = "Base Periodo Septiembre";
			ultraGridColumn70.Header.VisiblePosition = 77;
			ultraGridColumn70.Width = 180;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance55;
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.Caption = "Sueldo Periodo Septiembre";
			ultraGridColumn71.Header.VisiblePosition = 78;
			ultraGridColumn71.Width = 180;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance56;
			ultraGridColumn72.Format = "#,##0.00";
			ultraGridColumn72.Header.Caption = "Sobre Sueldo Periodo Septiembre";
			ultraGridColumn72.Header.VisiblePosition = 79;
			ultraGridColumn72.Width = 180;
			ultraGridColumn73.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance57;
			ultraGridColumn73.Format = "#,##0.00";
			ultraGridColumn73.Header.Caption = "Aporte Personal Septiembre";
			ultraGridColumn73.Header.VisiblePosition = 80;
			ultraGridColumn73.Width = 180;
			ultraGridColumn74.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance58;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.Caption = "FR Acum Septiembre";
			ultraGridColumn74.Header.VisiblePosition = 81;
			ultraGridColumn74.Width = 180;
			ultraGridColumn75.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance59;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.Caption = "FR No Acum Septiembre";
			ultraGridColumn75.Header.VisiblePosition = 82;
			ultraGridColumn75.Width = 180;
			ultraGridColumn76.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance60;
			ultraGridColumn76.Format = "#,##0.00";
			ultraGridColumn76.Header.Caption = "Base Periodo Octubre";
			ultraGridColumn76.Header.VisiblePosition = 85;
			ultraGridColumn76.Width = 180;
			ultraGridColumn77.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance61;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.Caption = "Sueldo Periodo Octubre";
			ultraGridColumn77.Header.VisiblePosition = 86;
			ultraGridColumn77.Width = 180;
			ultraGridColumn78.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance62;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.Caption = "Sobre Sueldo Periodo Octubre";
			ultraGridColumn78.Header.VisiblePosition = 87;
			ultraGridColumn78.Width = 180;
			ultraGridColumn79.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance63;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.Caption = "Aporte Personal Octubre";
			ultraGridColumn79.Header.VisiblePosition = 88;
			ultraGridColumn79.Width = 180;
			ultraGridColumn80.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance64;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.Caption = "FR Acum Octubre";
			ultraGridColumn80.Header.VisiblePosition = 89;
			ultraGridColumn80.Width = 180;
			ultraGridColumn81.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn81.CellAppearance = appearance65;
			ultraGridColumn81.Format = "#,##0.00";
			ultraGridColumn81.Header.Caption = "FR No Acum Octubre";
			ultraGridColumn81.Header.VisiblePosition = 90;
			ultraGridColumn81.Width = 180;
			ultraGridColumn82.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn82.CellAppearance = appearance66;
			ultraGridColumn82.Format = "#,##0.00";
			ultraGridColumn82.Header.Caption = "Base Periodo Noviembre";
			ultraGridColumn82.Header.VisiblePosition = 93;
			ultraGridColumn82.Width = 180;
			ultraGridColumn83.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn83.CellAppearance = appearance67;
			ultraGridColumn83.Format = "#,##0.00";
			ultraGridColumn83.Header.Caption = "Sueldo Periodo Noviembre";
			ultraGridColumn83.Header.VisiblePosition = 94;
			ultraGridColumn83.Width = 180;
			ultraGridColumn84.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn84.CellAppearance = appearance68;
			ultraGridColumn84.Format = "#,##0.00";
			ultraGridColumn84.Header.Caption = "Sobre Sueldo Periodo Noviembre";
			ultraGridColumn84.Header.VisiblePosition = 95;
			ultraGridColumn84.Width = 180;
			ultraGridColumn85.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance69.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn85.CellAppearance = appearance69;
			ultraGridColumn85.Format = "#,##0.00";
			ultraGridColumn85.Header.Caption = "Aporte Personall Noviembre";
			ultraGridColumn85.Header.VisiblePosition = 96;
			ultraGridColumn85.Width = 180;
			ultraGridColumn86.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn86.CellAppearance = appearance70;
			ultraGridColumn86.Format = "#,##0.00";
			ultraGridColumn86.Header.Caption = "FR Acum Noviembre";
			ultraGridColumn86.Header.VisiblePosition = 97;
			ultraGridColumn86.Width = 180;
			ultraGridColumn87.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn87.CellAppearance = appearance71;
			ultraGridColumn87.Format = "#,##0.00";
			ultraGridColumn87.Header.Caption = "FR No Acum Noviembre";
			ultraGridColumn87.Header.VisiblePosition = 98;
			ultraGridColumn87.Width = 180;
			ultraGridColumn88.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn88.CellAppearance = appearance72;
			ultraGridColumn88.Format = "#,##0.00";
			ultraGridColumn88.Header.VisiblePosition = 19;
			ultraGridColumn89.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance73.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn89.CellAppearance = appearance73;
			ultraGridColumn89.Format = "#,##0.00";
			ultraGridColumn89.Header.VisiblePosition = 20;
			ultraGridColumn90.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn90.CellAppearance = appearance74;
			ultraGridColumn90.Format = "#,##0.00";
			ultraGridColumn90.Header.VisiblePosition = 27;
			ultraGridColumn91.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn91.CellAppearance = appearance75;
			ultraGridColumn91.Format = "#,##0.00";
			ultraGridColumn91.Header.VisiblePosition = 28;
			ultraGridColumn92.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn92.CellAppearance = appearance76;
			ultraGridColumn92.Format = "#,##0.00";
			ultraGridColumn92.Header.VisiblePosition = 35;
			ultraGridColumn93.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn93.CellAppearance = appearance77;
			ultraGridColumn93.Format = "#,##0.00";
			ultraGridColumn93.Header.VisiblePosition = 36;
			ultraGridColumn94.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn94.CellAppearance = appearance78;
			ultraGridColumn94.Format = "#,##0.00";
			ultraGridColumn94.Header.VisiblePosition = 43;
			ultraGridColumn95.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn95.CellAppearance = appearance79;
			ultraGridColumn95.Format = "#,##0.00";
			ultraGridColumn95.Header.VisiblePosition = 44;
			ultraGridColumn96.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn96.CellAppearance = appearance80;
			ultraGridColumn96.Format = "#,##0.00";
			ultraGridColumn96.Header.VisiblePosition = 51;
			ultraGridColumn97.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn97.CellAppearance = appearance81;
			ultraGridColumn97.Format = "#,##0.00";
			ultraGridColumn97.Header.VisiblePosition = 52;
			ultraGridColumn98.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn98.CellAppearance = appearance82;
			ultraGridColumn98.Format = "#,##0.00";
			ultraGridColumn98.Header.VisiblePosition = 59;
			ultraGridColumn99.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn99.CellAppearance = appearance83;
			ultraGridColumn99.Format = "#,##0.00";
			ultraGridColumn99.Header.VisiblePosition = 60;
			ultraGridColumn100.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn100.CellAppearance = appearance84;
			ultraGridColumn100.Format = "#,##0.00";
			ultraGridColumn100.Header.VisiblePosition = 67;
			ultraGridColumn101.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn101.CellAppearance = appearance85;
			ultraGridColumn101.Format = "#,##0.00";
			ultraGridColumn101.Header.VisiblePosition = 68;
			ultraGridColumn102.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance86.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn102.CellAppearance = appearance86;
			ultraGridColumn102.Format = "#,##0.00";
			ultraGridColumn102.Header.VisiblePosition = 75;
			ultraGridColumn103.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance87.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn103.CellAppearance = appearance87;
			ultraGridColumn103.Format = "#,##0.00";
			ultraGridColumn103.Header.VisiblePosition = 76;
			ultraGridColumn104.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance88.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn104.CellAppearance = appearance88;
			ultraGridColumn104.Format = "#,##0.00";
			ultraGridColumn104.Header.VisiblePosition = 83;
			ultraGridColumn105.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance89.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn105.CellAppearance = appearance89;
			ultraGridColumn105.Format = "#,##0.00";
			ultraGridColumn105.Header.VisiblePosition = 84;
			ultraGridColumn106.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance90.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn106.CellAppearance = appearance90;
			ultraGridColumn106.Format = "#,##0.00";
			ultraGridColumn106.Header.VisiblePosition = 91;
			ultraGridColumn107.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance91.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn107.CellAppearance = appearance91;
			ultraGridColumn107.Format = "#,##0.00";
			ultraGridColumn107.Header.VisiblePosition = 92;
			ultraGridColumn108.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance92.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn108.CellAppearance = appearance92;
			ultraGridColumn108.Format = "#,##0.00";
			ultraGridColumn108.Header.VisiblePosition = 99;
			ultraGridColumn109.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance93.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn109.CellAppearance = appearance93;
			ultraGridColumn109.Format = "#,##0.00";
			ultraGridColumn109.Header.VisiblePosition = 100;
			ultraGridColumn110.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn110.CellAppearance = appearance94;
			ultraGridColumn110.Format = "#,##0.00";
			ultraGridColumn110.Header.VisiblePosition = 107;
			ultraGridColumn111.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn111.CellAppearance = appearance95;
			ultraGridColumn111.Format = "#,##0.00";
			ultraGridColumn111.Header.VisiblePosition = 108;
			ultraGridColumn112.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance96.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn112.CellAppearance = appearance96;
			ultraGridColumn112.Format = "#,##0.00";
			ultraGridColumn112.Header.Caption = "Base Periodo Diciembre";
			ultraGridColumn112.Header.VisiblePosition = 101;
			ultraGridColumn113.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance97.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn113.CellAppearance = appearance97;
			ultraGridColumn113.Format = "#,##0.00";
			ultraGridColumn113.Header.Caption = "Sueldo Periodo Diciembre";
			ultraGridColumn113.Header.VisiblePosition = 102;
			ultraGridColumn114.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance98.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn114.CellAppearance = appearance98;
			ultraGridColumn114.Format = "#,##0.00";
			ultraGridColumn114.Header.Caption = "Sobre Sueldo Periodo Diciembre";
			ultraGridColumn114.Header.VisiblePosition = 103;
			ultraGridColumn115.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance99.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn115.CellAppearance = appearance99;
			ultraGridColumn115.Format = "#,##0.00";
			ultraGridColumn115.Header.Caption = "Aporte Personal Diciembre";
			ultraGridColumn115.Header.VisiblePosition = 104;
			ultraGridColumn116.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance100.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn116.CellAppearance = appearance100;
			ultraGridColumn116.Format = "#,##0.00";
			ultraGridColumn116.Header.Caption = "FR Acum Diciembre";
			ultraGridColumn116.Header.VisiblePosition = 105;
			ultraGridColumn117.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance101.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn117.CellAppearance = appearance101;
			ultraGridColumn117.Format = "#,##0.00";
			ultraGridColumn117.Header.Caption = "FR No Acum Diciembre";
			ultraGridColumn117.Header.VisiblePosition = 106;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn117});
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.SummaryFooterCaption = "Totales";
			this.uGridRoles.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGridRoles.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridRoles.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridRoles.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance102.BackColor = System.Drawing.Color.Transparent;
			this.uGridRoles.DisplayLayout.Override.CardAreaAppearance = appearance102;
			appearance103.ForeColor = System.Drawing.Color.Black;
			appearance103.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridRoles.DisplayLayout.Override.CellAppearance = appearance103;
			appearance104.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance104.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance104.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance104.FontData.BoldAsString = "True";
			appearance104.FontData.Name = "Arial";
			appearance104.FontData.SizeInPoints = 8.5F;
			appearance104.ForeColor = System.Drawing.Color.White;
			appearance104.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRoles.DisplayLayout.Override.HeaderAppearance = appearance104;
			this.uGridRoles.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance105.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance105.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance105.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRoles.DisplayLayout.Override.RowSelectorAppearance = appearance105;
			appearance106.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance106.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance106.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRoles.DisplayLayout.Override.SelectedRowAppearance = appearance106;
			this.uGridRoles.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRoles.Location = new System.Drawing.Point(8, 120);
			this.uGridRoles.Name = "uGridRoles";
			this.uGridRoles.Size = new System.Drawing.Size(1232, 256);
			this.uGridRoles.TabIndex = 792;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn35.DataType = typeof(System.Decimal);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn47.DataType = typeof(System.Decimal);
			ultraDataColumn48.DataType = typeof(System.Decimal);
			ultraDataColumn49.DataType = typeof(System.Decimal);
			ultraDataColumn50.DataType = typeof(System.Decimal);
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn53.DataType = typeof(System.Decimal);
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn55.DataType = typeof(System.Decimal);
			ultraDataColumn56.DataType = typeof(System.Decimal);
			ultraDataColumn57.DataType = typeof(System.Decimal);
			ultraDataColumn58.DataType = typeof(System.Decimal);
			ultraDataColumn59.DataType = typeof(System.Decimal);
			ultraDataColumn60.DataType = typeof(System.Decimal);
			ultraDataColumn61.DataType = typeof(System.Decimal);
			ultraDataColumn62.DataType = typeof(System.Decimal);
			ultraDataColumn63.DataType = typeof(System.Decimal);
			ultraDataColumn64.DataType = typeof(System.Decimal);
			ultraDataColumn65.DataType = typeof(System.Decimal);
			ultraDataColumn66.DataType = typeof(System.Decimal);
			ultraDataColumn67.DataType = typeof(System.Decimal);
			ultraDataColumn68.DataType = typeof(System.Decimal);
			ultraDataColumn69.DataType = typeof(System.Decimal);
			ultraDataColumn70.DataType = typeof(System.Decimal);
			ultraDataColumn71.DataType = typeof(System.Decimal);
			ultraDataColumn72.DataType = typeof(System.Decimal);
			ultraDataColumn73.DataType = typeof(System.Decimal);
			ultraDataColumn74.DataType = typeof(System.Decimal);
			ultraDataColumn75.DataType = typeof(System.Decimal);
			ultraDataColumn76.DataType = typeof(System.Decimal);
			ultraDataColumn77.DataType = typeof(System.Decimal);
			ultraDataColumn78.DataType = typeof(System.Decimal);
			ultraDataColumn79.DataType = typeof(System.Decimal);
			ultraDataColumn80.DataType = typeof(System.Decimal);
			ultraDataColumn81.DataType = typeof(System.Decimal);
			ultraDataColumn82.DataType = typeof(System.Decimal);
			ultraDataColumn83.DataType = typeof(System.Decimal);
			ultraDataColumn84.DataType = typeof(System.Decimal);
			ultraDataColumn85.DataType = typeof(System.Decimal);
			ultraDataColumn86.DataType = typeof(System.Decimal);
			ultraDataColumn87.DataType = typeof(System.Decimal);
			ultraDataColumn88.DataType = typeof(System.Decimal);
			ultraDataColumn89.DataType = typeof(System.Decimal);
			ultraDataColumn90.DataType = typeof(System.Decimal);
			ultraDataColumn91.DataType = typeof(System.Decimal);
			ultraDataColumn92.DataType = typeof(System.Decimal);
			ultraDataColumn93.DataType = typeof(System.Decimal);
			ultraDataColumn94.DataType = typeof(System.Decimal);
			ultraDataColumn95.DataType = typeof(System.Decimal);
			ultraDataColumn96.DataType = typeof(System.Decimal);
			ultraDataColumn97.DataType = typeof(System.Decimal);
			ultraDataColumn98.DataType = typeof(System.Decimal);
			ultraDataColumn99.DataType = typeof(System.Decimal);
			ultraDataColumn100.DataType = typeof(System.Decimal);
			ultraDataColumn101.DataType = typeof(System.Decimal);
			ultraDataColumn102.DataType = typeof(System.Decimal);
			ultraDataColumn103.DataType = typeof(System.Decimal);
			ultraDataColumn104.DataType = typeof(System.Decimal);
			ultraDataColumn105.DataType = typeof(System.Decimal);
			ultraDataColumn106.DataType = typeof(System.Decimal);
			ultraDataColumn107.DataType = typeof(System.Decimal);
			ultraDataColumn108.DataType = typeof(System.Decimal);
			ultraDataColumn109.DataType = typeof(System.Decimal);
			ultraDataColumn110.DataType = typeof(System.Decimal);
			ultraDataColumn111.DataType = typeof(System.Decimal);
			ultraDataColumn112.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn112});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1464, 8);
			this.groupBox1.TabIndex = 791;
			this.groupBox1.TabStop = false;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1048, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 794;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1144, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(88, 23);
			this.btnSalir.TabIndex = 793;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn113.DataType = typeof(bool);
			ultraDataColumn113.DefaultValue = false;
			ultraDataColumn114.DataType = typeof(int);
			ultraDataColumn115.DataType = typeof(int);
			ultraDataColumn116.DataType = typeof(int);
			ultraDataColumn117.DataType = typeof(int);
			ultraDataColumn121.DataType = typeof(System.DateTime);
			ultraDataColumn122.DataType = typeof(System.DateTime);
			ultraDataColumn123.DataType = typeof(int);
			ultraDataColumn124.DataType = typeof(int);
			ultraDataColumn125.DataType = typeof(int);
			ultraDataColumn126.DataType = typeof(int);
			ultraDataColumn127.DataType = typeof(int);
			ultraDataColumn129.DataType = typeof(int);
			ultraDataColumn134.DataType = typeof(System.Decimal);
			ultraDataColumn134.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn135.DataType = typeof(System.Decimal);
			ultraDataColumn135.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn136.DataType = typeof(System.Decimal);
			ultraDataColumn136.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn137.DataType = typeof(System.Decimal);
			ultraDataColumn137.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn138.DataType = typeof(System.Decimal);
			ultraDataColumn138.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn139.DataType = typeof(System.Decimal);
			ultraDataColumn139.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn140.DataType = typeof(System.Decimal);
			ultraDataColumn140.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn141.DataType = typeof(System.Decimal);
			ultraDataColumn141.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn142.DataType = typeof(int);
			ultraDataColumn143.DataType = typeof(System.Decimal);
			ultraDataColumn143.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn144.DataType = typeof(System.Decimal);
			ultraDataColumn144.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn145.DataType = typeof(int);
			ultraDataColumn146.DataType = typeof(System.Decimal);
			ultraDataColumn146.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn147.DataType = typeof(System.Decimal);
			ultraDataColumn147.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn148.DataType = typeof(int);
			ultraDataColumn149.DataType = typeof(int);
			ultraDataColumn150.DataType = typeof(System.Decimal);
			ultraDataColumn150.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn151.DataType = typeof(System.Decimal);
			ultraDataColumn151.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn152.DataType = typeof(System.Decimal);
			ultraDataColumn152.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn153.DataType = typeof(System.Decimal);
			ultraDataColumn153.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn154.DataType = typeof(int);
			ultraDataColumn155.DataType = typeof(System.Decimal);
			ultraDataColumn155.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn156.DataType = typeof(System.Decimal);
			ultraDataColumn156.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn157.DataType = typeof(System.Decimal);
			ultraDataColumn157.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn158.DataType = typeof(System.Decimal);
			ultraDataColumn158.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn159.DataType = typeof(System.Decimal);
			ultraDataColumn159.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn160.DataType = typeof(System.Decimal);
			ultraDataColumn161.DataType = typeof(System.Decimal);
			ultraDataColumn161.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn162.DataType = typeof(System.Decimal);
			ultraDataColumn162.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn163.DataType = typeof(System.Decimal);
			ultraDataColumn163.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn164.DataType = typeof(System.Decimal);
			ultraDataColumn164.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn165.DataType = typeof(System.Decimal);
			ultraDataColumn165.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn166.DataType = typeof(System.Decimal);
			ultraDataColumn166.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn167.DataType = typeof(System.Decimal);
			ultraDataColumn167.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn168.DataType = typeof(System.Decimal);
			ultraDataColumn168.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn169.DataType = typeof(System.Decimal);
			ultraDataColumn169.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn170.DataType = typeof(System.Decimal);
			ultraDataColumn170.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn171.DataType = typeof(System.Decimal);
			ultraDataColumn171.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn172.DataType = typeof(System.Decimal);
			ultraDataColumn172.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn173.DataType = typeof(System.Decimal);
			ultraDataColumn173.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn174.DataType = typeof(System.Decimal);
			ultraDataColumn174.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn175.DataType = typeof(System.Decimal);
			ultraDataColumn175.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn176.DataType = typeof(System.Decimal);
			ultraDataColumn176.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn177.DataType = typeof(bool);
			ultraDataColumn177.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn146,
																																 ultraDataColumn147,
																																 ultraDataColumn148,
																																 ultraDataColumn149,
																																 ultraDataColumn150,
																																 ultraDataColumn151,
																																 ultraDataColumn152,
																																 ultraDataColumn153,
																																 ultraDataColumn154,
																																 ultraDataColumn155,
																																 ultraDataColumn156,
																																 ultraDataColumn157,
																																 ultraDataColumn158,
																																 ultraDataColumn159,
																																 ultraDataColumn160,
																																 ultraDataColumn161,
																																 ultraDataColumn162,
																																 ultraDataColumn163,
																																 ultraDataColumn164,
																																 ultraDataColumn165,
																																 ultraDataColumn166,
																																 ultraDataColumn167,
																																 ultraDataColumn168,
																																 ultraDataColumn169,
																																 ultraDataColumn170,
																																 ultraDataColumn171,
																																 ultraDataColumn172,
																																 ultraDataColumn173,
																																 ultraDataColumn174,
																																 ultraDataColumn175,
																																 ultraDataColumn176,
																																 ultraDataColumn177,
																																 ultraDataColumn178,
																																 ultraDataColumn179,
																																 ultraDataColumn180,
																																 ultraDataColumn181});
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
			// frmReporteRDEP
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1248, 382);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.uGridRoles);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtIdRolNomina);
			this.Controls.Add(this.cmbNEDepartamento);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label33);
			this.Controls.Add(this.cmbNECargo);
			this.Controls.Add(this.txtAnio);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.cmbNECentroCosto);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.KeyPreview = true;
			this.Name = "frmReporteRDEP";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "RDEP";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteRDEP_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteRDEP_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolNomina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRoles)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.Close();
		}
		int Estado = 0;
		private Acceso miAcceso;
	

		private void frmReporteRDEP_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "050712");
			
			if (!Funcion.Permiso("1853", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a RDEP", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				this.txtAnio.MinValue = 2019;
				this.txtAnio.MaxValue = DateTime.Today.Year;
				this.txtAnio.Value = DateTime.Today.Year;
				this.cmbNECentroCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 45, 0");
				this.cmbNEDepartamento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 48, 0");
				this.cmbNECargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 49, 0");
				if (miAcceso.Buscar) 
				{
					this.btnVer.Enabled = true;
				
				}
				if (miAcceso.BExportar) this.btnExportar.Enabled = true;
			}
		}

		private void txtAnio_ValueChanged(object sender, System.EventArgs e)
		{
			string sSQLNo = string.Format("Exec [ReporteRDEP] {0}, {1}, {2}, {3}, '{4}'", 5,5,5,0,"");
			this.uGridRoles.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLNo);
			this.lblContador.Text ="";
		}

		private void cmbNECentroCosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
						if(e.KeyCode == Keys.Enter) this.cmbNEDepartamento.Focus();
		}

		private void cmbNEDepartamento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNECargo.Focus();
		}

		private void cmbNECargo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int idProyecto = 0;
			int idDepartamento = 0;
			int idCargo = 0;

			if (this.cmbNECentroCosto.ActiveRow != null) idProyecto = (int)this.cmbNECentroCosto.Value;
			if (this.cmbNEDepartamento.ActiveRow != null) idDepartamento = (int)this.cmbNEDepartamento.Value;
			if (this.cmbNECargo.ActiveRow != null) idCargo = (int)this.cmbNECargo.Value;

			string sSQL = string.Format("Exec [ReporteRDEP] {0}, {1}, {2}, {3}, '{4}'", this.txtAnio.Value,idDepartamento,idCargo,idProyecto, this.txtBuscar.Text.ToString());
			this.uGridRoles.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.lblContador.Text = this.uGridRoles.Rows.Count + " REGISTROS ENCONTRADOS";		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}

		private void frmReporteRDEP_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
					FuncionesProcedimientos.ExportaExcel(this.uGridRoles);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}
	}
}
