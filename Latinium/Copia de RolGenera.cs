using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolGenera.
	/// </summary>
	public class RolGenera : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private C1.Data.C1DataSet cdsRolGrupo;
		private C1.Data.C1DataView cdvRolPago;
		private C1.Data.C1DataView cdvRolConcepto;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdRolPago;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.Misc.UltraButton btGuardar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDepartamento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSucursal;
		private C1.Data.C1DataSet cdsRolTabla;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSucursal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private Infragistics.Win.Misc.UltraButton btContabilizar;
		private Infragistics.Win.Misc.UltraButton btEgreso;
		private Infragistics.Win.Misc.UltraButton btEditar;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoGrupo;
		private C1.Data.C1DataView cdvSucursal;
		private Infragistics.Win.Misc.UltraButton btEnvio;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRolGrupo;
		private Infragistics.Win.Misc.UltraButton btRubro;
		private Infragistics.Win.Misc.UltraButton btPersonal;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.ComponentModel.IContainer components;

		public RolGenera()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Basico");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasLabor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Faltas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FactorTiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("anio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("mes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoIII");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vacacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FondoRes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoIV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso1");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso2");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso3");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso4");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso5");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso6");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso7");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso8");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso9");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso10");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso11");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso12");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso13");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso14");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso15");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso1");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso2");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso3");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso4");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso5");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso6");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso7");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso8");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso9");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso10");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso11");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso12");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso13");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso14");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso15");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provision1");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provision2");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provision3");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provision4");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provision5");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provision6");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provision7");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provision8");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provision9");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provision10");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor1");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor2");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor3");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor4");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor5");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor6");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor7");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor8");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor9");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor10");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor11");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor12");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor13");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor14");
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor15");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoProv1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoProv2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoProv3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoProv4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoProv5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Utilidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso16");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso17");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso18");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso19");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso20");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso21");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso22");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso23");
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso24");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso25");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso26");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso27");
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso28");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso29");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso30");
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso31");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso32");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso33");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso34");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso35");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso36");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso37");
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso38");
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso39");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso40");
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso41");
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso42");
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso43");
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso44");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso45");
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso46");
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso47");
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso48");
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso49");
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ingreso50");
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso16");
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso17");
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso18");
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso19");
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso20");
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso21");
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso22");
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso23");
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso24");
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso25");
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso26");
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso27");
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso28");
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso29");
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso30");
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso31");
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso32");
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso33");
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso34");
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso35");
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso36");
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso37");
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso38");
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso39");
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso40");
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso41");
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso42");
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso43");
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso44");
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso45");
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso46");
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso47");
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso48");
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso49");
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("egreso50");
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor16");
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor17");
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor18");
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor19");
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor20");
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor21");
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor22");
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor23");
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor24");
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor25");
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor26");
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor27");
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor28");
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor29");
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor30");
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor31");
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor32");
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor33");
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor34");
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor35");
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor36");
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor37");
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor38");
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor39");
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor40");
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor41");
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor42");
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor43");
			Infragistics.Win.Appearance appearance155 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor44");
			Infragistics.Win.Appearance appearance156 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor45");
			Infragistics.Win.Appearance appearance157 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor46");
			Infragistics.Win.Appearance appearance158 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor47");
			Infragistics.Win.Appearance appearance159 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor48");
			Infragistics.Win.Appearance appearance160 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor49");
			Infragistics.Win.Appearance appearance161 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor50");
			Infragistics.Win.Appearance appearance162 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoProv6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoProv7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoProv8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoProv9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoProv10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso1", 23, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso1", 23, true);
			Infragistics.Win.Appearance appearance163 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso2", 24, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso2", 24, true);
			Infragistics.Win.Appearance appearance164 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso3", 25, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso3", 25, true);
			Infragistics.Win.Appearance appearance165 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso4", 26, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso4", 26, true);
			Infragistics.Win.Appearance appearance166 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso5", 27, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso5", 27, true);
			Infragistics.Win.Appearance appearance167 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso6", 28, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso6", 28, true);
			Infragistics.Win.Appearance appearance168 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso7", 29, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso7", 29, true);
			Infragistics.Win.Appearance appearance169 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso8", 30, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso8", 30, true);
			Infragistics.Win.Appearance appearance170 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso9", 31, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso9", 31, true);
			Infragistics.Win.Appearance appearance171 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso10", 32, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso10", 32, true);
			Infragistics.Win.Appearance appearance172 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso11", 33, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso11", 33, true);
			Infragistics.Win.Appearance appearance173 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso12", 34, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso12", 34, true);
			Infragistics.Win.Appearance appearance174 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso13", 35, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso13", 35, true);
			Infragistics.Win.Appearance appearance175 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso14", 36, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso14", 36, true);
			Infragistics.Win.Appearance appearance176 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso15", 37, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso15", 37, true);
			Infragistics.Win.Appearance appearance177 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso1", 38, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso1", 38, true);
			Infragistics.Win.Appearance appearance178 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso2", 39, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso2", 39, true);
			Infragistics.Win.Appearance appearance179 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso3", 40, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso3", 40, true);
			Infragistics.Win.Appearance appearance180 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso4", 41, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso4", 41, true);
			Infragistics.Win.Appearance appearance181 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso5", 42, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso5", 42, true);
			Infragistics.Win.Appearance appearance182 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso6", 43, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso6", 43, true);
			Infragistics.Win.Appearance appearance183 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso7", 44, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso7", 44, true);
			Infragistics.Win.Appearance appearance184 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso8", 45, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso8", 45, true);
			Infragistics.Win.Appearance appearance185 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso9", 46, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso9", 46, true);
			Infragistics.Win.Appearance appearance186 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso10", 47, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso10", 47, true);
			Infragistics.Win.Appearance appearance187 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso11", 48, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso11", 48, true);
			Infragistics.Win.Appearance appearance188 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings27 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso12", 49, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso12", 49, true);
			Infragistics.Win.Appearance appearance189 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings28 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso13", 50, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso13", 50, true);
			Infragistics.Win.Appearance appearance190 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings29 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso14", 51, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso14", 51, true);
			Infragistics.Win.Appearance appearance191 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings30 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso15", 52, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso15", 52, true);
			Infragistics.Win.Appearance appearance192 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings31 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "provision1", 53, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "provision1", 53, true);
			Infragistics.Win.Appearance appearance193 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings32 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "provision2", 54, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "provision2", 54, true);
			Infragistics.Win.Appearance appearance194 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings33 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "provision3", 55, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "provision3", 55, true);
			Infragistics.Win.Appearance appearance195 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings34 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "provision4", 56, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "provision4", 56, true);
			Infragistics.Win.Appearance appearance196 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings35 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "provision5", 57, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "provision5", 57, true);
			Infragistics.Win.Appearance appearance197 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings36 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "provision6", 58, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "provision6", 58, true);
			Infragistics.Win.Appearance appearance198 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings37 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "provision7", 59, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "provision7", 59, true);
			Infragistics.Win.Appearance appearance199 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings38 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "provision8", 60, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "provision8", 60, true);
			Infragistics.Win.Appearance appearance200 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings39 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor1", 63, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor1", 63, true);
			Infragistics.Win.Appearance appearance201 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings40 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor2", 64, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor2", 64, true);
			Infragistics.Win.Appearance appearance202 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings41 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor3", 65, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor3", 65, true);
			Infragistics.Win.Appearance appearance203 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings42 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor4", 66, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor4", 66, true);
			Infragistics.Win.Appearance appearance204 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings43 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor5", 67, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor5", 67, true);
			Infragistics.Win.Appearance appearance205 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings44 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor6", 68, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor6", 68, true);
			Infragistics.Win.Appearance appearance206 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings45 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor7", 69, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor7", 69, true);
			Infragistics.Win.Appearance appearance207 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings46 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor8", 70, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor8", 70, true);
			Infragistics.Win.Appearance appearance208 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings47 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor9", 71, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor9", 71, true);
			Infragistics.Win.Appearance appearance209 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings48 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor10", 72, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor10", 72, true);
			Infragistics.Win.Appearance appearance210 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings49 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor11", 73, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor11", 73, true);
			Infragistics.Win.Appearance appearance211 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings50 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor12", 74, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor12", 74, true);
			Infragistics.Win.Appearance appearance212 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings51 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor13", 75, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor13", 75, true);
			Infragistics.Win.Appearance appearance213 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings52 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor14", 76, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor14", 76, true);
			Infragistics.Win.Appearance appearance214 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings53 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor15", 77, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor15", 77, true);
			Infragistics.Win.Appearance appearance215 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings54 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "provision9", 61, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "provision9", 61, true);
			Infragistics.Win.Appearance appearance216 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings55 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "provision10", 62, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "provision10", 62, true);
			Infragistics.Win.Appearance appearance217 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings56 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso16", 92, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso16", 92, true);
			Infragistics.Win.Appearance appearance218 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings57 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso17", 93, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso17", 93, true);
			Infragistics.Win.Appearance appearance219 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings58 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso18", 94, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso18", 94, true);
			Infragistics.Win.Appearance appearance220 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings59 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso19", 95, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso19", 95, true);
			Infragistics.Win.Appearance appearance221 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings60 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso20", 96, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso20", 96, true);
			Infragistics.Win.Appearance appearance222 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings61 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso21", 97, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso21", 97, true);
			Infragistics.Win.Appearance appearance223 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings62 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso22", 98, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso22", 98, true);
			Infragistics.Win.Appearance appearance224 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings63 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso23", 99, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso23", 99, true);
			Infragistics.Win.Appearance appearance225 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings64 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso24", 100, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso24", 100, true);
			Infragistics.Win.Appearance appearance226 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings65 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso25", 101, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso25", 101, true);
			Infragistics.Win.Appearance appearance227 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings66 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso26", 102, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso26", 102, true);
			Infragistics.Win.Appearance appearance228 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings67 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso27", 103, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso27", 103, true);
			Infragistics.Win.Appearance appearance229 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings68 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso28", 104, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso28", 104, true);
			Infragistics.Win.Appearance appearance230 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings69 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso29", 105, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso29", 105, true);
			Infragistics.Win.Appearance appearance231 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings70 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso30", 106, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso30", 106, true);
			Infragistics.Win.Appearance appearance232 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings71 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso31", 107, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso31", 107, true);
			Infragistics.Win.Appearance appearance233 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings72 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso32", 108, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso32", 108, true);
			Infragistics.Win.Appearance appearance234 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings73 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso33", 109, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso33", 109, true);
			Infragistics.Win.Appearance appearance235 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings74 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso34", 110, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso34", 110, true);
			Infragistics.Win.Appearance appearance236 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings75 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso35", 111, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso35", 111, true);
			Infragistics.Win.Appearance appearance237 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings76 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso36", 112, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso36", 112, true);
			Infragistics.Win.Appearance appearance238 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings77 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso37", 113, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso37", 113, true);
			Infragistics.Win.Appearance appearance239 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings78 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso38", 114, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso38", 114, true);
			Infragistics.Win.Appearance appearance240 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings79 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso39", 115, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso39", 115, true);
			Infragistics.Win.Appearance appearance241 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings80 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso40", 116, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso40", 116, true);
			Infragistics.Win.Appearance appearance242 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings81 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso41", 117, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso41", 117, true);
			Infragistics.Win.Appearance appearance243 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings82 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso42", 118, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso42", 118, true);
			Infragistics.Win.Appearance appearance244 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings83 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso43", 119, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso43", 119, true);
			Infragistics.Win.Appearance appearance245 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings84 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso44", 120, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso44", 120, true);
			Infragistics.Win.Appearance appearance246 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings85 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso45", 121, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso45", 121, true);
			Infragistics.Win.Appearance appearance247 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings86 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso46", 122, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso46", 122, true);
			Infragistics.Win.Appearance appearance248 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings87 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso47", 123, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso47", 123, true);
			Infragistics.Win.Appearance appearance249 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings88 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso48", 124, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso48", 124, true);
			Infragistics.Win.Appearance appearance250 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings89 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso49", 125, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso49", 125, true);
			Infragistics.Win.Appearance appearance251 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings90 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ingreso50", 126, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ingreso50", 126, true);
			Infragistics.Win.Appearance appearance252 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings91 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso16", 127, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso16", 127, true);
			Infragistics.Win.Appearance appearance253 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings92 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso17", 128, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso17", 128, true);
			Infragistics.Win.Appearance appearance254 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings93 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso18", 129, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso18", 129, true);
			Infragistics.Win.Appearance appearance255 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings94 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso19", 130, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso19", 130, true);
			Infragistics.Win.Appearance appearance256 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings95 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso20", 131, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso20", 131, true);
			Infragistics.Win.Appearance appearance257 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings96 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso21", 132, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso21", 132, true);
			Infragistics.Win.Appearance appearance258 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings97 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso22", 133, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso22", 133, true);
			Infragistics.Win.Appearance appearance259 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings98 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso23", 134, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso23", 134, true);
			Infragistics.Win.Appearance appearance260 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings99 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso24", 135, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso24", 135, true);
			Infragistics.Win.Appearance appearance261 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings100 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso25", 136, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso25", 136, true);
			Infragistics.Win.Appearance appearance262 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings101 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso26", 137, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso26", 137, true);
			Infragistics.Win.Appearance appearance263 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings102 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso27", 138, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso27", 138, true);
			Infragistics.Win.Appearance appearance264 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings103 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso28", 139, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso28", 139, true);
			Infragistics.Win.Appearance appearance265 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings104 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso29", 140, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso29", 140, true);
			Infragistics.Win.Appearance appearance266 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings105 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso30", 141, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso30", 141, true);
			Infragistics.Win.Appearance appearance267 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings106 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso31", 142, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso31", 142, true);
			Infragistics.Win.Appearance appearance268 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings107 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso32", 143, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso32", 143, true);
			Infragistics.Win.Appearance appearance269 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings108 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso33", 144, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso33", 144, true);
			Infragistics.Win.Appearance appearance270 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings109 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso34", 145, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso34", 145, true);
			Infragistics.Win.Appearance appearance271 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings110 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso35", 146, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso35", 146, true);
			Infragistics.Win.Appearance appearance272 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings111 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso36", 147, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso36", 147, true);
			Infragistics.Win.Appearance appearance273 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings112 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso37", 148, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso37", 148, true);
			Infragistics.Win.Appearance appearance274 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings113 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso38", 149, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso38", 149, true);
			Infragistics.Win.Appearance appearance275 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings114 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso39", 150, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso39", 150, true);
			Infragistics.Win.Appearance appearance276 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings115 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Count, null, "Nombre", 14, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nombre", 14, true);
			Infragistics.Win.Appearance appearance277 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings116 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor16", 162, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor16", 162, true);
			Infragistics.Win.Appearance appearance278 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings117 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor17", 163, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor17", 163, true);
			Infragistics.Win.Appearance appearance279 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings118 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor18", 164, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor18", 164, true);
			Infragistics.Win.Appearance appearance280 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings119 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor19", 165, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor19", 165, true);
			Infragistics.Win.Appearance appearance281 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings120 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor20", 166, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor20", 166, true);
			Infragistics.Win.Appearance appearance282 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings121 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor21", 167, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor21", 167, true);
			Infragistics.Win.Appearance appearance283 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings122 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor22", 168, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor22", 168, true);
			Infragistics.Win.Appearance appearance284 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings123 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor23", 169, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor23", 169, true);
			Infragistics.Win.Appearance appearance285 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings124 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor24", 170, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor24", 170, true);
			Infragistics.Win.Appearance appearance286 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings125 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor25", 171, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor25", 171, true);
			Infragistics.Win.Appearance appearance287 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings126 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor26", 172, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor26", 172, true);
			Infragistics.Win.Appearance appearance288 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings127 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor27", 173, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor27", 173, true);
			Infragistics.Win.Appearance appearance289 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings128 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor28", 174, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor28", 174, true);
			Infragistics.Win.Appearance appearance290 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings129 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor29", 175, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor29", 175, true);
			Infragistics.Win.Appearance appearance291 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings130 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor30", 176, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor30", 176, true);
			Infragistics.Win.Appearance appearance292 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings131 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor31", 177, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor31", 177, true);
			Infragistics.Win.Appearance appearance293 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings132 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor32", 178, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor32", 178, true);
			Infragistics.Win.Appearance appearance294 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings133 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor33", 179, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor33", 179, true);
			Infragistics.Win.Appearance appearance295 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings134 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor34", 180, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor34", 180, true);
			Infragistics.Win.Appearance appearance296 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings135 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor35", 181, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor35", 181, true);
			Infragistics.Win.Appearance appearance297 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings136 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor36", 182, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor36", 182, true);
			Infragistics.Win.Appearance appearance298 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings137 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor37", 183, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor37", 183, true);
			Infragistics.Win.Appearance appearance299 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings138 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor38", 184, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor38", 184, true);
			Infragistics.Win.Appearance appearance300 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings139 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor39", 185, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor39", 185, true);
			Infragistics.Win.Appearance appearance301 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings140 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valor40", 186, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valor40", 186, true);
			Infragistics.Win.Appearance appearance302 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings141 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso41", 152, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso41", 152, true);
			Infragistics.Win.Appearance appearance303 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings142 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso42", 153, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso42", 153, true);
			Infragistics.Win.Appearance appearance304 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings143 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso43", 154, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso43", 154, true);
			Infragistics.Win.Appearance appearance305 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings144 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso44", 155, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso44", 155, true);
			Infragistics.Win.Appearance appearance306 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings145 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso45", 156, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso45", 156, true);
			Infragistics.Win.Appearance appearance307 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings146 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "egreso40", 151, true, "RolPago", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "egreso40", 151, true);
			Infragistics.Win.Appearance appearance308 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance309 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance310 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance311 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance312 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance313 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolSucursal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn205 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn206 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn207 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OrigenFondos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn208 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSeguro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn209 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipoEmpleador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn210 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn211 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn212 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolDepartamento", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn213 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn214 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn215 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn216 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn217 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.btGuardar = new Infragistics.Win.Misc.UltraButton();
			this.cdsRolGrupo = new C1.Data.C1DataSet();
			this.grdRolPago = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdvRolPago = new C1.Data.C1DataView();
			this.cdvRolConcepto = new C1.Data.C1DataView();
			this.chkDepartamento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbSucursal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvSucursal = new C1.Data.C1DataView();
			this.cdsRolTabla = new C1.Data.C1DataSet();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkSucursal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btContabilizar = new Infragistics.Win.Misc.UltraButton();
			this.btEgreso = new Infragistics.Win.Misc.UltraButton();
			this.btEditar = new Infragistics.Win.Misc.UltraButton();
			this.cmbRolGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.chkTodoGrupo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btEnvio = new Infragistics.Win.Misc.UltraButton();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btRubro = new Infragistics.Win.Misc.UltraButton();
			this.btPersonal = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRolPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRolGrupo)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbFecha
			// 
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(56, 9);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(104, 21);
			this.cmbFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 0;
			this.cmbFecha.Value = new System.DateTime(2006, 11, 5, 0, 0, 0, 0);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(16, 9);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(28, 15);
			this.ultraLabel1.TabIndex = 1;
			this.ultraLabel1.Text = "Mes:";
			// 
			// btGenerar
			// 
			this.btGenerar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGenerar.Location = new System.Drawing.Point(8, 34);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(56, 25);
			this.btGenerar.TabIndex = 2;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// btGuardar
			// 
			this.btGuardar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGuardar.Location = new System.Drawing.Point(136, 34);
			this.btGuardar.Name = "btGuardar";
			this.btGuardar.Size = new System.Drawing.Size(56, 25);
			this.btGuardar.TabIndex = 3;
			this.btGuardar.Text = "Guardar";
			this.btGuardar.Click += new System.EventHandler(this.btGuardar_Click);
			// 
			// cdsRolGrupo
			// 
			this.cdsRolGrupo.BindingContextCtrl = this;
			this.cdsRolGrupo.DataLibrary = "LibPersonal";
			this.cdsRolGrupo.DataLibraryUrl = "";
			this.cdsRolGrupo.DataSetDef = "dsRolGrupo";
			this.cdsRolGrupo.FillOnRequest = false;
			this.cdsRolGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolGrupo.Name = "cdsRolGrupo";
			this.cdsRolGrupo.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolGrupo.SchemaDef = null;
			this.cdsRolGrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolGrupo_BeforeFill);
			// 
			// grdRolPago
			// 
			this.grdRolPago.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdRolPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdRolPago.DataMember = "RolPago";
			this.grdRolPago.DataSource = this.cdsRolGrupo;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdRolPago.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 5;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 6;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 8;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 9;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 69;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 70;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 71;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 72;
			ultraGridColumn9.Hidden = true;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance2;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 68;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 66;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn12.Header.Caption = "Dias Labor.";
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 36;
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Width = 45;
			ultraGridColumn14.Header.VisiblePosition = 67;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Width = 159;
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 73;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn18.Header.Caption = "Año";
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Width = 43;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn19.Header.Caption = "Mes";
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Width = 37;
			ultraGridColumn20.Header.VisiblePosition = 74;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 75;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 76;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 77;
			ultraGridColumn23.Hidden = true;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance3;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 11;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance4;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 12;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance5;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 13;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance6;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 14;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance7;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 15;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance8;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 16;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance9;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 17;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance10;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 18;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance11;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.VisiblePosition = 19;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance12;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 20;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance13;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 21;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance14;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 22;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance15;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.VisiblePosition = 23;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance16;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 24;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance17;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.VisiblePosition = 25;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance18;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 26;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance19;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.VisiblePosition = 27;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance20;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 28;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance21;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.VisiblePosition = 29;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance22;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.VisiblePosition = 30;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance23;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 31;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance24;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.VisiblePosition = 32;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance25;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 33;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance26;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.VisiblePosition = 34;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance27;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.VisiblePosition = 35;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance28;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.VisiblePosition = 36;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance29;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.VisiblePosition = 37;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance30;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.VisiblePosition = 38;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance31;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.VisiblePosition = 39;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance32;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.VisiblePosition = 40;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance33;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.VisiblePosition = 41;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance34;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.VisiblePosition = 42;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance35;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.VisiblePosition = 43;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn57.CellAppearance = appearance36;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.VisiblePosition = 44;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance37;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.VisiblePosition = 45;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance38;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.VisiblePosition = 46;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance39;
			ultraGridColumn60.Format = "#,##0.00";
			ultraGridColumn60.Header.VisiblePosition = 47;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance40;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.VisiblePosition = 48;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance41;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.VisiblePosition = 49;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance42;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.VisiblePosition = 50;
			ultraGridColumn63.Width = 92;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance43;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.VisiblePosition = 51;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance44;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.VisiblePosition = 52;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance45;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.VisiblePosition = 53;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance46;
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.VisiblePosition = 54;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance47;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.VisiblePosition = 55;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance48;
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.VisiblePosition = 56;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance49;
			ultraGridColumn70.Format = "#,##0.00";
			ultraGridColumn70.Header.VisiblePosition = 57;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance50;
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.VisiblePosition = 58;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance51;
			ultraGridColumn72.Format = "#,##0.00";
			ultraGridColumn72.Header.VisiblePosition = 59;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance52;
			ultraGridColumn73.Format = "#,##0.00";
			ultraGridColumn73.Header.VisiblePosition = 60;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance53;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.VisiblePosition = 61;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance54;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.VisiblePosition = 62;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance55;
			ultraGridColumn76.Format = "#,##0.00";
			ultraGridColumn76.Header.VisiblePosition = 63;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance56;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.VisiblePosition = 64;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance57;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.VisiblePosition = 65;
			ultraGridColumn79.Header.VisiblePosition = 78;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 79;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 80;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 81;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 82;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 83;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 84;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 85;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 86;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 87;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 88;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 89;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 90;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 91;
			ultraGridColumn92.Hidden = true;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn93.CellAppearance = appearance58;
			ultraGridColumn93.Format = "#,##0.00";
			ultraGridColumn93.Header.VisiblePosition = 92;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn94.CellAppearance = appearance59;
			ultraGridColumn94.Format = "#,##0.00";
			ultraGridColumn94.Header.VisiblePosition = 93;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn95.CellAppearance = appearance60;
			ultraGridColumn95.Format = "#,##0.00";
			ultraGridColumn95.Header.VisiblePosition = 94;
			ultraGridColumn95.TabStop = false;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn96.CellAppearance = appearance61;
			ultraGridColumn96.Format = "#,##0.00";
			ultraGridColumn96.Header.VisiblePosition = 95;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn97.CellAppearance = appearance62;
			ultraGridColumn97.Format = "#,##0.00";
			ultraGridColumn97.Header.VisiblePosition = 96;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn98.CellAppearance = appearance63;
			ultraGridColumn98.Format = "#,##0.00";
			ultraGridColumn98.Header.VisiblePosition = 97;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn99.CellAppearance = appearance64;
			ultraGridColumn99.Format = "#,##0.00";
			ultraGridColumn99.Header.VisiblePosition = 98;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn100.CellAppearance = appearance65;
			ultraGridColumn100.Format = "#,##0.00";
			ultraGridColumn100.Header.VisiblePosition = 99;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn101.CellAppearance = appearance66;
			ultraGridColumn101.Format = "#,##0.00";
			ultraGridColumn101.Header.VisiblePosition = 100;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn102.CellAppearance = appearance67;
			ultraGridColumn102.Format = "#,##0.00";
			ultraGridColumn102.Header.VisiblePosition = 101;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn103.CellAppearance = appearance68;
			ultraGridColumn103.Format = "#,##0.00";
			ultraGridColumn103.Header.VisiblePosition = 102;
			appearance69.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn104.CellAppearance = appearance69;
			ultraGridColumn104.Format = "#,##0.00";
			ultraGridColumn104.Header.VisiblePosition = 103;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn105.CellAppearance = appearance70;
			ultraGridColumn105.Format = "#,##0.00";
			ultraGridColumn105.Header.VisiblePosition = 104;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn106.CellAppearance = appearance71;
			ultraGridColumn106.Format = "#,##0.00";
			ultraGridColumn106.Header.VisiblePosition = 105;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn107.CellAppearance = appearance72;
			ultraGridColumn107.Format = "#,##0.00";
			ultraGridColumn107.Header.VisiblePosition = 106;
			appearance73.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn108.CellAppearance = appearance73;
			ultraGridColumn108.Format = "#,##0.00";
			ultraGridColumn108.Header.VisiblePosition = 107;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn109.CellAppearance = appearance74;
			ultraGridColumn109.Format = "#,##0.00";
			ultraGridColumn109.Header.VisiblePosition = 108;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn110.CellAppearance = appearance75;
			ultraGridColumn110.Format = "#,##0.00";
			ultraGridColumn110.Header.VisiblePosition = 109;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn111.CellAppearance = appearance76;
			ultraGridColumn111.Format = "#,##0.00";
			ultraGridColumn111.Header.VisiblePosition = 110;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn112.CellAppearance = appearance77;
			ultraGridColumn112.Format = "#,##0.00";
			ultraGridColumn112.Header.VisiblePosition = 111;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn113.CellAppearance = appearance78;
			ultraGridColumn113.Format = "#,##0.00";
			ultraGridColumn113.Header.VisiblePosition = 112;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn114.CellAppearance = appearance79;
			ultraGridColumn114.Format = "#,##0.00";
			ultraGridColumn114.Header.VisiblePosition = 113;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn115.CellAppearance = appearance80;
			ultraGridColumn115.Format = "#,##0.00";
			ultraGridColumn115.Header.VisiblePosition = 114;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn116.CellAppearance = appearance81;
			ultraGridColumn116.Format = "#,##0.00";
			ultraGridColumn116.Header.VisiblePosition = 115;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn117.CellAppearance = appearance82;
			ultraGridColumn117.Format = "#,##0.00";
			ultraGridColumn117.Header.VisiblePosition = 116;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn118.CellAppearance = appearance83;
			ultraGridColumn118.Format = "#,##0.00";
			ultraGridColumn118.Header.VisiblePosition = 117;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn119.CellAppearance = appearance84;
			ultraGridColumn119.Format = "#,##0.00";
			ultraGridColumn119.Header.VisiblePosition = 118;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn120.CellAppearance = appearance85;
			ultraGridColumn120.Format = "#,##0.00";
			ultraGridColumn120.Header.VisiblePosition = 119;
			appearance86.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn121.CellAppearance = appearance86;
			ultraGridColumn121.Format = "#,##0.00";
			ultraGridColumn121.Header.VisiblePosition = 120;
			appearance87.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn122.CellAppearance = appearance87;
			ultraGridColumn122.Format = "#,##0.00";
			ultraGridColumn122.Header.VisiblePosition = 121;
			appearance88.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn123.CellAppearance = appearance88;
			ultraGridColumn123.Format = "#,##0.00";
			ultraGridColumn123.Header.VisiblePosition = 122;
			appearance89.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn124.CellAppearance = appearance89;
			ultraGridColumn124.Format = "#,##0.00";
			ultraGridColumn124.Header.VisiblePosition = 123;
			appearance90.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn125.CellAppearance = appearance90;
			ultraGridColumn125.Format = "#,##0.00";
			ultraGridColumn125.Header.VisiblePosition = 124;
			appearance91.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn126.CellAppearance = appearance91;
			ultraGridColumn126.Format = "#,##0.00";
			ultraGridColumn126.Header.VisiblePosition = 125;
			appearance92.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn127.CellAppearance = appearance92;
			ultraGridColumn127.Format = "#,##0.00";
			ultraGridColumn127.Header.VisiblePosition = 126;
			appearance93.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn128.CellAppearance = appearance93;
			ultraGridColumn128.Format = "#,##0.00";
			ultraGridColumn128.Header.VisiblePosition = 127;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn129.CellAppearance = appearance94;
			ultraGridColumn129.Format = "#,##0.00";
			ultraGridColumn129.Header.VisiblePosition = 128;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn130.CellAppearance = appearance95;
			ultraGridColumn130.Format = "#,##0.00";
			ultraGridColumn130.Header.VisiblePosition = 129;
			appearance96.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn131.CellAppearance = appearance96;
			ultraGridColumn131.Format = "#,##0.00";
			ultraGridColumn131.Header.VisiblePosition = 130;
			appearance97.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn132.CellAppearance = appearance97;
			ultraGridColumn132.Format = "#,##0.00";
			ultraGridColumn132.Header.VisiblePosition = 131;
			appearance98.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn133.CellAppearance = appearance98;
			ultraGridColumn133.Format = "#,##0.00";
			ultraGridColumn133.Header.VisiblePosition = 132;
			appearance99.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn134.CellAppearance = appearance99;
			ultraGridColumn134.Format = "#,##0.00";
			ultraGridColumn134.Header.VisiblePosition = 133;
			appearance100.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn135.CellAppearance = appearance100;
			ultraGridColumn135.Format = "#,##0.00";
			ultraGridColumn135.Header.VisiblePosition = 134;
			appearance101.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn136.CellAppearance = appearance101;
			ultraGridColumn136.Format = "#,##0.00";
			ultraGridColumn136.Header.VisiblePosition = 135;
			appearance102.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn137.CellAppearance = appearance102;
			ultraGridColumn137.Format = "#,##0.00";
			ultraGridColumn137.Header.VisiblePosition = 136;
			appearance103.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn138.CellAppearance = appearance103;
			ultraGridColumn138.Format = "#,##0.00";
			ultraGridColumn138.Header.VisiblePosition = 137;
			appearance104.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn139.CellAppearance = appearance104;
			ultraGridColumn139.Format = "#,##0.00";
			ultraGridColumn139.Header.VisiblePosition = 138;
			appearance105.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn140.CellAppearance = appearance105;
			ultraGridColumn140.Format = "#,##0.00";
			ultraGridColumn140.Header.VisiblePosition = 139;
			appearance106.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn141.CellAppearance = appearance106;
			ultraGridColumn141.Format = "#,##0.00";
			ultraGridColumn141.Header.VisiblePosition = 140;
			appearance107.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn142.CellAppearance = appearance107;
			ultraGridColumn142.Format = "#,##0.00";
			ultraGridColumn142.Header.VisiblePosition = 141;
			appearance108.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn143.CellAppearance = appearance108;
			ultraGridColumn143.Format = "#,##0.00";
			ultraGridColumn143.Header.VisiblePosition = 142;
			appearance109.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn144.CellAppearance = appearance109;
			ultraGridColumn144.Format = "#,##0.00";
			ultraGridColumn144.Header.VisiblePosition = 143;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn145.CellAppearance = appearance110;
			ultraGridColumn145.Format = "#,##0.00";
			ultraGridColumn145.Header.VisiblePosition = 144;
			appearance111.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn146.CellAppearance = appearance111;
			ultraGridColumn146.Format = "#,##0.00";
			ultraGridColumn146.Header.VisiblePosition = 145;
			appearance112.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn147.CellAppearance = appearance112;
			ultraGridColumn147.Format = "#,##0.00";
			ultraGridColumn147.Header.VisiblePosition = 146;
			appearance113.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn148.CellAppearance = appearance113;
			ultraGridColumn148.Format = "#,##0.00";
			ultraGridColumn148.Header.VisiblePosition = 147;
			appearance114.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn149.CellAppearance = appearance114;
			ultraGridColumn149.Format = "#,##0.00";
			ultraGridColumn149.Header.VisiblePosition = 148;
			appearance115.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn150.CellAppearance = appearance115;
			ultraGridColumn150.Format = "#,##0.00";
			ultraGridColumn150.Header.VisiblePosition = 149;
			appearance116.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn151.CellAppearance = appearance116;
			ultraGridColumn151.Format = "#,##0.00";
			ultraGridColumn151.Header.VisiblePosition = 150;
			appearance117.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn152.CellAppearance = appearance117;
			ultraGridColumn152.Format = "#,##0.00";
			ultraGridColumn152.Header.VisiblePosition = 151;
			appearance118.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn153.CellAppearance = appearance118;
			ultraGridColumn153.Format = "#,##0.00";
			ultraGridColumn153.Header.VisiblePosition = 152;
			appearance119.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn154.CellAppearance = appearance119;
			ultraGridColumn154.Format = "#,##0.00";
			ultraGridColumn154.Header.VisiblePosition = 153;
			appearance120.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn155.CellAppearance = appearance120;
			ultraGridColumn155.Format = "#,##0.00";
			ultraGridColumn155.Header.VisiblePosition = 154;
			appearance121.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn156.CellAppearance = appearance121;
			ultraGridColumn156.Format = "#,##0.00";
			ultraGridColumn156.Header.VisiblePosition = 155;
			appearance122.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn157.CellAppearance = appearance122;
			ultraGridColumn157.Format = "#,##0.00";
			ultraGridColumn157.Header.VisiblePosition = 156;
			appearance123.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn158.CellAppearance = appearance123;
			ultraGridColumn158.Format = "#,##0.00";
			ultraGridColumn158.Header.VisiblePosition = 157;
			appearance124.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn159.CellAppearance = appearance124;
			ultraGridColumn159.Format = "#,##0.00";
			ultraGridColumn159.Header.VisiblePosition = 158;
			appearance125.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn160.CellAppearance = appearance125;
			ultraGridColumn160.Format = "#,##0.00";
			ultraGridColumn160.Header.VisiblePosition = 159;
			appearance126.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn161.CellAppearance = appearance126;
			ultraGridColumn161.Format = "#,##0.00";
			ultraGridColumn161.Header.VisiblePosition = 160;
			appearance127.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn162.CellAppearance = appearance127;
			ultraGridColumn162.Format = "#,##0.00";
			ultraGridColumn162.Header.VisiblePosition = 161;
			appearance128.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn163.CellAppearance = appearance128;
			ultraGridColumn163.Format = "#,##0.00";
			ultraGridColumn163.Header.VisiblePosition = 162;
			appearance129.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn164.CellAppearance = appearance129;
			ultraGridColumn164.Format = "#,##0.00";
			ultraGridColumn164.Header.VisiblePosition = 163;
			appearance130.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn165.CellAppearance = appearance130;
			ultraGridColumn165.Format = "#,##0.00";
			ultraGridColumn165.Header.VisiblePosition = 164;
			appearance131.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn166.CellAppearance = appearance131;
			ultraGridColumn166.Format = "#,##0.00";
			ultraGridColumn166.Header.VisiblePosition = 165;
			appearance132.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn167.CellAppearance = appearance132;
			ultraGridColumn167.Format = "#,##0.00";
			ultraGridColumn167.Header.VisiblePosition = 166;
			appearance133.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn168.CellAppearance = appearance133;
			ultraGridColumn168.Format = "#,##0.00";
			ultraGridColumn168.Header.VisiblePosition = 167;
			appearance134.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn169.CellAppearance = appearance134;
			ultraGridColumn169.Format = "#,##0.00";
			ultraGridColumn169.Header.VisiblePosition = 168;
			appearance135.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn170.CellAppearance = appearance135;
			ultraGridColumn170.Format = "#,##0.00";
			ultraGridColumn170.Header.VisiblePosition = 169;
			appearance136.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn171.CellAppearance = appearance136;
			ultraGridColumn171.Format = "#,##0.00";
			ultraGridColumn171.Header.VisiblePosition = 170;
			appearance137.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn172.CellAppearance = appearance137;
			ultraGridColumn172.Format = "#,##0.00";
			ultraGridColumn172.Header.VisiblePosition = 171;
			appearance138.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn173.CellAppearance = appearance138;
			ultraGridColumn173.Format = "#,##0.00";
			ultraGridColumn173.Header.VisiblePosition = 172;
			appearance139.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn174.CellAppearance = appearance139;
			ultraGridColumn174.Format = "#,##0.00";
			ultraGridColumn174.Header.VisiblePosition = 173;
			appearance140.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn175.CellAppearance = appearance140;
			ultraGridColumn175.Format = "#,##0.00";
			ultraGridColumn175.Header.VisiblePosition = 174;
			appearance141.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn176.CellAppearance = appearance141;
			ultraGridColumn176.Format = "#,##0.00";
			ultraGridColumn176.Header.VisiblePosition = 175;
			appearance142.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn177.CellAppearance = appearance142;
			ultraGridColumn177.Format = "#,##0.00";
			ultraGridColumn177.Header.VisiblePosition = 176;
			appearance143.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn178.CellAppearance = appearance143;
			ultraGridColumn178.Format = "#,##0.00";
			ultraGridColumn178.Header.VisiblePosition = 177;
			appearance144.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn179.CellAppearance = appearance144;
			ultraGridColumn179.Format = "#,##0.00";
			ultraGridColumn179.Header.VisiblePosition = 178;
			appearance145.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn180.CellAppearance = appearance145;
			ultraGridColumn180.Format = "#,##0.00";
			ultraGridColumn180.Header.VisiblePosition = 179;
			appearance146.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn181.CellAppearance = appearance146;
			ultraGridColumn181.Format = "#,##0.00";
			ultraGridColumn181.Header.VisiblePosition = 180;
			appearance147.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn182.CellAppearance = appearance147;
			ultraGridColumn182.Format = "#,##0.00";
			ultraGridColumn182.Header.VisiblePosition = 181;
			appearance148.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn183.CellAppearance = appearance148;
			ultraGridColumn183.Format = "#,##0.00";
			ultraGridColumn183.Header.VisiblePosition = 182;
			appearance149.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn184.CellAppearance = appearance149;
			ultraGridColumn184.Format = "#,##0.00";
			ultraGridColumn184.Header.VisiblePosition = 183;
			appearance150.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn185.CellAppearance = appearance150;
			ultraGridColumn185.Format = "#,##0.00";
			ultraGridColumn185.Header.VisiblePosition = 184;
			appearance151.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn186.CellAppearance = appearance151;
			ultraGridColumn186.Format = "#,##0.00";
			ultraGridColumn186.Header.VisiblePosition = 185;
			appearance152.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn187.CellAppearance = appearance152;
			ultraGridColumn187.Format = "#,##0.00";
			ultraGridColumn187.Header.VisiblePosition = 186;
			appearance153.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn188.CellAppearance = appearance153;
			ultraGridColumn188.Format = "#,##0.00";
			ultraGridColumn188.Header.VisiblePosition = 187;
			appearance154.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn189.CellAppearance = appearance154;
			ultraGridColumn189.Format = "#,##0.00";
			ultraGridColumn189.Header.VisiblePosition = 188;
			appearance155.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn190.CellAppearance = appearance155;
			ultraGridColumn190.Format = "#,##0.00";
			ultraGridColumn190.Header.VisiblePosition = 189;
			appearance156.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn191.CellAppearance = appearance156;
			ultraGridColumn191.Format = "#,##0.00";
			ultraGridColumn191.Header.VisiblePosition = 190;
			appearance157.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn192.CellAppearance = appearance157;
			ultraGridColumn192.Format = "#,##0.00";
			ultraGridColumn192.Header.VisiblePosition = 191;
			appearance158.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn193.CellAppearance = appearance158;
			ultraGridColumn193.Format = "#,##0.00";
			ultraGridColumn193.Header.VisiblePosition = 192;
			appearance159.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn194.CellAppearance = appearance159;
			ultraGridColumn194.Format = "#,##0.00";
			ultraGridColumn194.Header.VisiblePosition = 193;
			appearance160.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn195.CellAppearance = appearance160;
			ultraGridColumn195.Format = "#,##0.00";
			ultraGridColumn195.Header.VisiblePosition = 194;
			appearance161.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn196.CellAppearance = appearance161;
			ultraGridColumn196.Format = "#,##0.00";
			ultraGridColumn196.Header.VisiblePosition = 195;
			appearance162.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn197.CellAppearance = appearance162;
			ultraGridColumn197.Format = "#,##0.00";
			ultraGridColumn197.Header.VisiblePosition = 196;
			ultraGridColumn198.Header.VisiblePosition = 197;
			ultraGridColumn198.Hidden = true;
			ultraGridColumn199.Header.VisiblePosition = 198;
			ultraGridColumn199.Hidden = true;
			ultraGridColumn200.Header.VisiblePosition = 199;
			ultraGridColumn200.Hidden = true;
			ultraGridColumn201.Header.VisiblePosition = 200;
			ultraGridColumn201.Hidden = true;
			ultraGridColumn202.Header.VisiblePosition = 201;
			ultraGridColumn202.Hidden = true;
			ultraGridColumn203.Header.VisiblePosition = 202;
			ultraGridColumn203.Hidden = true;
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
																										 ultraGridColumn146,
																										 ultraGridColumn147,
																										 ultraGridColumn148,
																										 ultraGridColumn149,
																										 ultraGridColumn150,
																										 ultraGridColumn151,
																										 ultraGridColumn152,
																										 ultraGridColumn153,
																										 ultraGridColumn154,
																										 ultraGridColumn155,
																										 ultraGridColumn156,
																										 ultraGridColumn157,
																										 ultraGridColumn158,
																										 ultraGridColumn159,
																										 ultraGridColumn160,
																										 ultraGridColumn161,
																										 ultraGridColumn162,
																										 ultraGridColumn163,
																										 ultraGridColumn164,
																										 ultraGridColumn165,
																										 ultraGridColumn166,
																										 ultraGridColumn167,
																										 ultraGridColumn168,
																										 ultraGridColumn169,
																										 ultraGridColumn170,
																										 ultraGridColumn171,
																										 ultraGridColumn172,
																										 ultraGridColumn173,
																										 ultraGridColumn174,
																										 ultraGridColumn175,
																										 ultraGridColumn176,
																										 ultraGridColumn177,
																										 ultraGridColumn178,
																										 ultraGridColumn179,
																										 ultraGridColumn180,
																										 ultraGridColumn181,
																										 ultraGridColumn182,
																										 ultraGridColumn183,
																										 ultraGridColumn184,
																										 ultraGridColumn185,
																										 ultraGridColumn186,
																										 ultraGridColumn187,
																										 ultraGridColumn188,
																										 ultraGridColumn189,
																										 ultraGridColumn190,
																										 ultraGridColumn191,
																										 ultraGridColumn192,
																										 ultraGridColumn193,
																										 ultraGridColumn194,
																										 ultraGridColumn195,
																										 ultraGridColumn196,
																										 ultraGridColumn197,
																										 ultraGridColumn198,
																										 ultraGridColumn199,
																										 ultraGridColumn200,
																										 ultraGridColumn201,
																										 ultraGridColumn202,
																										 ultraGridColumn203});
			appearance163.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance163;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance164.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance164;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance165.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance165;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance166.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance166;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance167.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance167;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance168.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance168;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance169.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance169;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance170.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance170;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance171.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance171;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance172.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance172;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance173.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance173;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance174.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance174;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance175.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance175;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance176.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance176;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance177.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance177;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance178.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance178;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance179.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance179;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance180.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance180;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance181.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance181;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance182.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance182;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance183.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance183;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance184.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance184;
			summarySettings22.DisplayFormat = "{0: #,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance185.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance185;
			summarySettings23.DisplayFormat = "{0: #,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance186.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance186;
			summarySettings24.DisplayFormat = "{0: #,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance187.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance187;
			summarySettings25.DisplayFormat = "{0: #,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance188.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance188;
			summarySettings26.DisplayFormat = "{0: #,##0.00}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance189.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings27.Appearance = appearance189;
			summarySettings27.DisplayFormat = "{0: #,##0.00}";
			summarySettings27.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance190.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings28.Appearance = appearance190;
			summarySettings28.DisplayFormat = "{0: #,##0.00}";
			summarySettings28.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance191.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings29.Appearance = appearance191;
			summarySettings29.DisplayFormat = "{0: #,##0.00}";
			summarySettings29.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance192.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings30.Appearance = appearance192;
			summarySettings30.DisplayFormat = "{0: #,##0.00}";
			summarySettings30.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance193.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings31.Appearance = appearance193;
			summarySettings31.DisplayFormat = "{0: #,##0.00}";
			summarySettings31.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance194.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings32.Appearance = appearance194;
			summarySettings32.DisplayFormat = "{0: #,##0.00}";
			summarySettings32.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance195.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings33.Appearance = appearance195;
			summarySettings33.DisplayFormat = "{0: #,##0.00}";
			summarySettings33.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance196.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings34.Appearance = appearance196;
			summarySettings34.DisplayFormat = "{0: #,##0.00}";
			summarySettings34.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance197.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings35.Appearance = appearance197;
			summarySettings35.DisplayFormat = "{0: #,##0.00}";
			summarySettings35.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance198.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings36.Appearance = appearance198;
			summarySettings36.DisplayFormat = "{0: #,##0.00}";
			summarySettings36.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance199.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings37.Appearance = appearance199;
			summarySettings37.DisplayFormat = "{0: #,##0.00}";
			summarySettings37.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance200.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings38.Appearance = appearance200;
			summarySettings38.DisplayFormat = "{0: #,##0.00}";
			summarySettings38.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance201.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings39.Appearance = appearance201;
			summarySettings39.DisplayFormat = "{0: #,##0.00}";
			summarySettings39.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance202.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings40.Appearance = appearance202;
			summarySettings40.DisplayFormat = "{0: #,##0.00}";
			summarySettings40.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance203.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings41.Appearance = appearance203;
			summarySettings41.DisplayFormat = "{0: #,##0.00}";
			summarySettings41.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance204.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings42.Appearance = appearance204;
			summarySettings42.DisplayFormat = "{0: #,##0.00}";
			summarySettings42.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance205.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings43.Appearance = appearance205;
			summarySettings43.DisplayFormat = "{0: #,##0.00}";
			summarySettings43.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance206.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings44.Appearance = appearance206;
			summarySettings44.DisplayFormat = "{0: #,##0.00}";
			summarySettings44.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance207.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings45.Appearance = appearance207;
			summarySettings45.DisplayFormat = "{0: #,##0.00}";
			summarySettings45.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance208.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings46.Appearance = appearance208;
			summarySettings46.DisplayFormat = "{0: #,##0.00}";
			summarySettings46.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance209.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings47.Appearance = appearance209;
			summarySettings47.DisplayFormat = "{0: #,##0.00}";
			summarySettings47.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance210.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings48.Appearance = appearance210;
			summarySettings48.DisplayFormat = "{0: #,##0.00}";
			summarySettings48.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance211.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings49.Appearance = appearance211;
			summarySettings49.DisplayFormat = "{0: #,##0.00}";
			summarySettings49.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance212.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings50.Appearance = appearance212;
			summarySettings50.DisplayFormat = "{0: #,##0.00}";
			summarySettings50.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance213.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings51.Appearance = appearance213;
			summarySettings51.DisplayFormat = "{0: #,##0.00}";
			summarySettings51.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance214.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings52.Appearance = appearance214;
			summarySettings52.DisplayFormat = "{0: #,##0.00}";
			summarySettings52.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance215.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings53.Appearance = appearance215;
			summarySettings53.DisplayFormat = "{0: #,##0.00}";
			summarySettings53.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance216.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings54.Appearance = appearance216;
			summarySettings54.DisplayFormat = "{0: #,##0.00}";
			summarySettings54.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance217.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings55.Appearance = appearance217;
			summarySettings55.DisplayFormat = "{0: #,##0.00}";
			summarySettings55.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance218.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings56.Appearance = appearance218;
			summarySettings56.DisplayFormat = "{0: #,##0.00}";
			summarySettings56.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance219.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings57.Appearance = appearance219;
			summarySettings57.DisplayFormat = "{0: #,##0.00}";
			summarySettings57.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance220.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings58.Appearance = appearance220;
			summarySettings58.DisplayFormat = "{0: #,##0.00}";
			summarySettings58.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance221.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings59.Appearance = appearance221;
			summarySettings59.DisplayFormat = "{0: #,##0.00}";
			summarySettings59.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance222.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings60.Appearance = appearance222;
			summarySettings60.DisplayFormat = "{0: #,##0.00}";
			summarySettings60.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance223.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings61.Appearance = appearance223;
			summarySettings61.DisplayFormat = "{0: #,##0.00}";
			summarySettings61.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance224.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings62.Appearance = appearance224;
			summarySettings62.DisplayFormat = "{0: #,##0.00}";
			summarySettings62.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance225.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings63.Appearance = appearance225;
			summarySettings63.DisplayFormat = "{0: #,##0.00}";
			summarySettings63.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance226.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings64.Appearance = appearance226;
			summarySettings64.DisplayFormat = "{0: #,##0.00}";
			summarySettings64.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance227.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings65.Appearance = appearance227;
			summarySettings65.DisplayFormat = "{0: #,##0.00}";
			summarySettings65.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance228.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings66.Appearance = appearance228;
			summarySettings66.DisplayFormat = "{0: #,##0.00}";
			summarySettings66.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance229.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings67.Appearance = appearance229;
			summarySettings67.DisplayFormat = "{0: #,##0.00}";
			summarySettings67.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance230.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings68.Appearance = appearance230;
			summarySettings68.DisplayFormat = "{0: #,##0.00}";
			summarySettings68.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance231.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings69.Appearance = appearance231;
			summarySettings69.DisplayFormat = "{0: #,##0.00}";
			summarySettings69.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance232.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings70.Appearance = appearance232;
			summarySettings70.DisplayFormat = "{0: #,##0.00}";
			summarySettings70.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance233.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings71.Appearance = appearance233;
			summarySettings71.DisplayFormat = "{0: #,##0.00}";
			summarySettings71.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance234.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings72.Appearance = appearance234;
			summarySettings72.DisplayFormat = "{0: #,##0.00}";
			summarySettings72.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance235.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings73.Appearance = appearance235;
			summarySettings73.DisplayFormat = "{0: #,##0.00}";
			summarySettings73.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance236.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings74.Appearance = appearance236;
			summarySettings74.DisplayFormat = "{0: #,##0.00}";
			summarySettings74.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance237.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings75.Appearance = appearance237;
			summarySettings75.DisplayFormat = "{0: #,##0.00}";
			summarySettings75.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance238.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings76.Appearance = appearance238;
			summarySettings76.DisplayFormat = "{0: #,##0.00}";
			summarySettings76.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance239.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings77.Appearance = appearance239;
			summarySettings77.DisplayFormat = "{0: #,##0.00}";
			summarySettings77.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance240.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings78.Appearance = appearance240;
			summarySettings78.DisplayFormat = "{0: #,##0.00}";
			summarySettings78.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance241.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings79.Appearance = appearance241;
			summarySettings79.DisplayFormat = "{0: #,##0.00}";
			summarySettings79.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance242.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings80.Appearance = appearance242;
			summarySettings80.DisplayFormat = "{0: #,##0.00}";
			summarySettings80.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance243.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings81.Appearance = appearance243;
			summarySettings81.DisplayFormat = "{0: #,##0.00}";
			summarySettings81.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance244.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings82.Appearance = appearance244;
			summarySettings82.DisplayFormat = "{0: #,##0.00}";
			summarySettings82.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance245.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings83.Appearance = appearance245;
			summarySettings83.DisplayFormat = "{0: #,##0.00}";
			summarySettings83.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance246.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings84.Appearance = appearance246;
			summarySettings84.DisplayFormat = "{0: #,##0.00}";
			summarySettings84.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance247.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings85.Appearance = appearance247;
			summarySettings85.DisplayFormat = "{0: #,##0.00}";
			summarySettings85.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance248.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings86.Appearance = appearance248;
			summarySettings86.DisplayFormat = "{0: #,##0.00}";
			summarySettings86.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance249.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings87.Appearance = appearance249;
			summarySettings87.DisplayFormat = "{0: #,##0.00}";
			summarySettings87.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance250.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings88.Appearance = appearance250;
			summarySettings88.DisplayFormat = "{0: #,##0.00}";
			summarySettings88.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance251.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings89.Appearance = appearance251;
			summarySettings89.DisplayFormat = "{0: #,##0.00}";
			summarySettings89.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance252.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings90.Appearance = appearance252;
			summarySettings90.DisplayFormat = "{0: #,##0.00}";
			summarySettings90.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance253.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings91.Appearance = appearance253;
			summarySettings91.DisplayFormat = "{0: #,##0.00}";
			summarySettings91.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance254.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings92.Appearance = appearance254;
			summarySettings92.DisplayFormat = "{0: #,##0.00}";
			summarySettings92.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance255.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings93.Appearance = appearance255;
			summarySettings93.DisplayFormat = "{0: #,##0.00}";
			summarySettings93.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance256.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings94.Appearance = appearance256;
			summarySettings94.DisplayFormat = "{0: #,##0.00}";
			summarySettings94.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance257.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings95.Appearance = appearance257;
			summarySettings95.DisplayFormat = "{0: #,##0.00}";
			summarySettings95.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance258.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings96.Appearance = appearance258;
			summarySettings96.DisplayFormat = "{0: #,##0.00}";
			summarySettings96.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance259.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings97.Appearance = appearance259;
			summarySettings97.DisplayFormat = "{0: #,##0.00}";
			summarySettings97.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance260.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings98.Appearance = appearance260;
			summarySettings98.DisplayFormat = "{0: #,##0.00}";
			summarySettings98.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance261.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings99.Appearance = appearance261;
			summarySettings99.DisplayFormat = "{0: #,##0.00}";
			summarySettings99.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance262.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings100.Appearance = appearance262;
			summarySettings100.DisplayFormat = "{0: #,##0.00}";
			summarySettings100.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance263.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings101.Appearance = appearance263;
			summarySettings101.DisplayFormat = "{0: #,##0.00}";
			summarySettings101.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance264.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings102.Appearance = appearance264;
			summarySettings102.DisplayFormat = "{0: #,##0.00}";
			summarySettings102.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance265.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings103.Appearance = appearance265;
			summarySettings103.DisplayFormat = "{0: #,##0.00}";
			summarySettings103.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance266.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings104.Appearance = appearance266;
			summarySettings104.DisplayFormat = "{0: #,##0.00}";
			summarySettings104.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance267.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings105.Appearance = appearance267;
			summarySettings105.DisplayFormat = "{0: #,##0.00}";
			summarySettings105.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance268.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings106.Appearance = appearance268;
			summarySettings106.DisplayFormat = "{0: #,##0.00}";
			summarySettings106.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance269.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings107.Appearance = appearance269;
			summarySettings107.DisplayFormat = "{0: #,##0.00}";
			summarySettings107.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance270.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings108.Appearance = appearance270;
			summarySettings108.DisplayFormat = "{0: #,##0.00}";
			summarySettings108.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance271.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings109.Appearance = appearance271;
			summarySettings109.DisplayFormat = "{0: #,##0.00}";
			summarySettings109.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance272.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings110.Appearance = appearance272;
			summarySettings110.DisplayFormat = "{0: #,##0.00}";
			summarySettings110.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance273.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings111.Appearance = appearance273;
			summarySettings111.DisplayFormat = "{0: #,##0.00}";
			summarySettings111.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance274.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings112.Appearance = appearance274;
			summarySettings112.DisplayFormat = "{0: #,##0.00}";
			summarySettings112.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance275.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings113.Appearance = appearance275;
			summarySettings113.DisplayFormat = "{0: #,##0.00}";
			summarySettings113.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance276.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings114.Appearance = appearance276;
			summarySettings114.DisplayFormat = "{0: #,##0.00}";
			summarySettings114.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance277.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings115.Appearance = appearance277;
			summarySettings115.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance278.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings116.Appearance = appearance278;
			summarySettings116.DisplayFormat = "{0: #,##0.00}";
			summarySettings116.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance279.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings117.Appearance = appearance279;
			summarySettings117.DisplayFormat = "{0: #,##0.00}";
			summarySettings117.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance280.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings118.Appearance = appearance280;
			summarySettings118.DisplayFormat = "{0: #,##0.00}";
			summarySettings118.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance281.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings119.Appearance = appearance281;
			summarySettings119.DisplayFormat = "{0: #,##0.00}";
			summarySettings119.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance282.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings120.Appearance = appearance282;
			summarySettings120.DisplayFormat = "{0: #,##0.00}";
			summarySettings120.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance283.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings121.Appearance = appearance283;
			summarySettings121.DisplayFormat = "{0: #,##0.00}";
			summarySettings121.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance284.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings122.Appearance = appearance284;
			summarySettings122.DisplayFormat = "{0: #,##0.00}";
			summarySettings122.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance285.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings123.Appearance = appearance285;
			summarySettings123.DisplayFormat = "{0: #,##0.00}";
			summarySettings123.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance286.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings124.Appearance = appearance286;
			summarySettings124.DisplayFormat = "{0: #,##0.00}";
			summarySettings124.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance287.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings125.Appearance = appearance287;
			summarySettings125.DisplayFormat = "{0: #,##0.00}";
			summarySettings125.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance288.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings126.Appearance = appearance288;
			summarySettings126.DisplayFormat = "{0: #,##0.00}";
			summarySettings126.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance289.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings127.Appearance = appearance289;
			summarySettings127.DisplayFormat = "{0: #,##0.00}";
			summarySettings127.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance290.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings128.Appearance = appearance290;
			summarySettings128.DisplayFormat = "{0: #,##0.00}";
			summarySettings128.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance291.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings129.Appearance = appearance291;
			summarySettings129.DisplayFormat = "{0: #,##0.00}";
			summarySettings129.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance292.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings130.Appearance = appearance292;
			summarySettings130.DisplayFormat = "{0: #,##0.00}";
			summarySettings130.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance293.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings131.Appearance = appearance293;
			summarySettings131.DisplayFormat = "{0: #,##0.00}";
			summarySettings131.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance294.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings132.Appearance = appearance294;
			summarySettings132.DisplayFormat = "{0: #,##0.00}";
			summarySettings132.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance295.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings133.Appearance = appearance295;
			summarySettings133.DisplayFormat = "{0: #,##0.00}";
			summarySettings133.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance296.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings134.Appearance = appearance296;
			summarySettings134.DisplayFormat = "{0: #,##0.00}";
			summarySettings134.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance297.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings135.Appearance = appearance297;
			summarySettings135.DisplayFormat = "{0: #,##0.00}";
			summarySettings135.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance298.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings136.Appearance = appearance298;
			summarySettings136.DisplayFormat = "{0: #,##0.00}";
			summarySettings136.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance299.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings137.Appearance = appearance299;
			summarySettings137.DisplayFormat = "{0: #,##0.00}";
			summarySettings137.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance300.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings138.Appearance = appearance300;
			summarySettings138.DisplayFormat = "{0: #,##0.00}";
			summarySettings138.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance301.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings139.Appearance = appearance301;
			summarySettings139.DisplayFormat = "{0: #,##0.00}";
			summarySettings139.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance302.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings140.Appearance = appearance302;
			summarySettings140.DisplayFormat = "{0: #,##0.00}";
			summarySettings140.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance303.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings141.Appearance = appearance303;
			summarySettings141.DisplayFormat = "{0: #,##0.00}";
			summarySettings141.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance304.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings142.Appearance = appearance304;
			summarySettings142.DisplayFormat = "{0: #,##0.00}";
			summarySettings142.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance305.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings143.Appearance = appearance305;
			summarySettings143.DisplayFormat = "{0: #,##0.00}";
			summarySettings143.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance306.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings144.Appearance = appearance306;
			summarySettings144.DisplayFormat = "{0: #,##0.00}";
			summarySettings144.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance307.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings145.Appearance = appearance307;
			summarySettings145.DisplayFormat = "{0: #,##0.00}";
			summarySettings145.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance308.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings146.Appearance = appearance308;
			summarySettings146.DisplayFormat = "{0: #,##0.00}";
			summarySettings146.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
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
																																															summarySettings67,
																																															summarySettings68,
																																															summarySettings69,
																																															summarySettings70,
																																															summarySettings71,
																																															summarySettings72,
																																															summarySettings73,
																																															summarySettings74,
																																															summarySettings75,
																																															summarySettings76,
																																															summarySettings77,
																																															summarySettings78,
																																															summarySettings79,
																																															summarySettings80,
																																															summarySettings81,
																																															summarySettings82,
																																															summarySettings83,
																																															summarySettings84,
																																															summarySettings85,
																																															summarySettings86,
																																															summarySettings87,
																																															summarySettings88,
																																															summarySettings89,
																																															summarySettings90,
																																															summarySettings91,
																																															summarySettings92,
																																															summarySettings93,
																																															summarySettings94,
																																															summarySettings95,
																																															summarySettings96,
																																															summarySettings97,
																																															summarySettings98,
																																															summarySettings99,
																																															summarySettings100,
																																															summarySettings101,
																																															summarySettings102,
																																															summarySettings103,
																																															summarySettings104,
																																															summarySettings105,
																																															summarySettings106,
																																															summarySettings107,
																																															summarySettings108,
																																															summarySettings109,
																																															summarySettings110,
																																															summarySettings111,
																																															summarySettings112,
																																															summarySettings113,
																																															summarySettings114,
																																															summarySettings115,
																																															summarySettings116,
																																															summarySettings117,
																																															summarySettings118,
																																															summarySettings119,
																																															summarySettings120,
																																															summarySettings121,
																																															summarySettings122,
																																															summarySettings123,
																																															summarySettings124,
																																															summarySettings125,
																																															summarySettings126,
																																															summarySettings127,
																																															summarySettings128,
																																															summarySettings129,
																																															summarySettings130,
																																															summarySettings131,
																																															summarySettings132,
																																															summarySettings133,
																																															summarySettings134,
																																															summarySettings135,
																																															summarySettings136,
																																															summarySettings137,
																																															summarySettings138,
																																															summarySettings139,
																																															summarySettings140,
																																															summarySettings141,
																																															summarySettings142,
																																															summarySettings143,
																																															summarySettings144,
																																															summarySettings145,
																																															summarySettings146});
			ultraGridBand1.SummaryFooterCaption = "";
			this.grdRolPago.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance309.BackColor = System.Drawing.Color.Transparent;
			this.grdRolPago.DisplayLayout.Override.CardAreaAppearance = appearance309;
			appearance310.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance310.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance310.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance310.FontData.BoldAsString = "True";
			appearance310.FontData.Name = "Arial";
			appearance310.FontData.SizeInPoints = 10F;
			appearance310.ForeColor = System.Drawing.Color.White;
			appearance310.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdRolPago.DisplayLayout.Override.HeaderAppearance = appearance310;
			this.grdRolPago.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance311.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance311.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance311.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRolPago.DisplayLayout.Override.RowSelectorAppearance = appearance311;
			appearance312.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance312.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance312.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRolPago.DisplayLayout.Override.SelectedRowAppearance = appearance312;
			appearance313.TextHAlign = Infragistics.Win.HAlign.Right;
			this.grdRolPago.DisplayLayout.Override.SummaryFooterAppearance = appearance313;
			this.grdRolPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdRolPago.Location = new System.Drawing.Point(8, 78);
			this.grdRolPago.Name = "grdRolPago";
			this.grdRolPago.Size = new System.Drawing.Size(904, 354);
			this.grdRolPago.TabIndex = 4;
			this.grdRolPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdRolPago_KeyDown);
			this.grdRolPago.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cdvRolPago
			// 
			this.cdvRolPago.BindingContextCtrl = this;
			this.cdvRolPago.DataSet = this.cdsRolGrupo;
			this.cdvRolPago.TableName = "";
			this.cdvRolPago.TableViewName = "RolPago";
			// 
			// cdvRolConcepto
			// 
			this.cdvRolConcepto.BindingContextCtrl = this;
			this.cdvRolConcepto.DataSet = this.cdsRolGrupo;
			this.cdvRolConcepto.TableName = "";
			this.cdvRolConcepto.TableViewName = "RolConcepto";
			// 
			// chkDepartamento
			// 
			this.chkDepartamento.Checked = true;
			this.chkDepartamento.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDepartamento.Location = new System.Drawing.Point(208, 9);
			this.chkDepartamento.Name = "chkDepartamento";
			this.chkDepartamento.Size = new System.Drawing.Size(136, 21);
			this.chkDepartamento.TabIndex = 6;
			this.chkDepartamento.Text = "Todo Departamento";
			// 
			// cmbSucursal
			// 
			this.cmbSucursal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSucursal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSucursal.DataSource = this.cdvSucursal;
			this.cmbSucursal.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn204.Header.VisiblePosition = 1;
			ultraGridColumn204.Hidden = true;
			ultraGridColumn204.Width = 19;
			ultraGridColumn205.Header.VisiblePosition = 0;
			ultraGridColumn205.Width = 197;
			ultraGridColumn206.Header.VisiblePosition = 2;
			ultraGridColumn206.Hidden = true;
			ultraGridColumn206.Width = 13;
			ultraGridColumn207.Header.VisiblePosition = 3;
			ultraGridColumn207.Hidden = true;
			ultraGridColumn207.Width = 22;
			ultraGridColumn208.Header.VisiblePosition = 4;
			ultraGridColumn208.Hidden = true;
			ultraGridColumn208.Width = 22;
			ultraGridColumn209.Header.VisiblePosition = 5;
			ultraGridColumn209.Hidden = true;
			ultraGridColumn209.Width = 26;
			ultraGridColumn210.Header.VisiblePosition = 6;
			ultraGridColumn210.Hidden = true;
			ultraGridColumn210.Width = 24;
			ultraGridColumn211.Header.VisiblePosition = 7;
			ultraGridColumn211.Hidden = true;
			ultraGridColumn211.Width = 22;
			ultraGridColumn212.Header.VisiblePosition = 8;
			ultraGridColumn212.Hidden = true;
			ultraGridColumn212.Width = 27;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn204,
																										 ultraGridColumn205,
																										 ultraGridColumn206,
																										 ultraGridColumn207,
																										 ultraGridColumn208,
																										 ultraGridColumn209,
																										 ultraGridColumn210,
																										 ultraGridColumn211,
																										 ultraGridColumn212});
			this.cmbSucursal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSucursal.DisplayMember = "Sucursal";
			this.cmbSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSucursal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSucursal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSucursal.Location = new System.Drawing.Point(352, 34);
			this.cmbSucursal.Name = "cmbSucursal";
			this.cmbSucursal.Size = new System.Drawing.Size(216, 21);
			this.cmbSucursal.TabIndex = 7;
			this.cmbSucursal.ValueMember = "idRolSucursal";
			this.cmbSucursal.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbSucursal_InitializeLayout);
			// 
			// cdvSucursal
			// 
			this.cdvSucursal.BindingContextCtrl = this;
			this.cdvSucursal.DataSet = this.cdsRolTabla;
			this.cdvSucursal.Sort = "Sucursal";
			this.cdvSucursal.TableName = "";
			this.cdvSucursal.TableViewName = "RolSucursal";
			// 
			// cdsRolTabla
			// 
			this.cdsRolTabla.BindingContextCtrl = this;
			this.cdsRolTabla.DataLibrary = "LibPersonal";
			this.cdsRolTabla.DataLibraryUrl = "";
			this.cdsRolTabla.DataSetDef = "dsRolTablas";
			this.cdsRolTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolTabla.Name = "cdsRolTabla";
			this.cdsRolTabla.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolTabla.SchemaDef = null;
			this.cdsRolTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolTabla_BeforeFill);
			// 
			// cmbDepartamento
			// 
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento.DataMember = "RolDepartamento";
			this.cmbDepartamento.DataSource = this.cdsRolTabla;
			this.cmbDepartamento.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn213.Header.VisiblePosition = 2;
			ultraGridColumn213.Hidden = true;
			ultraGridColumn214.Header.VisiblePosition = 0;
			ultraGridColumn214.Width = 197;
			ultraGridColumn215.Header.VisiblePosition = 1;
			ultraGridColumn215.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn213,
																										 ultraGridColumn214,
																										 ultraGridColumn215});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(352, 9);
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(216, 21);
			this.cmbDepartamento.TabIndex = 9;
			this.cmbDepartamento.ValueMember = "idDepartamento";
			// 
			// chkSucursal
			// 
			this.chkSucursal.Checked = true;
			this.chkSucursal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkSucursal.Location = new System.Drawing.Point(208, 34);
			this.chkSucursal.Name = "chkSucursal";
			this.chkSucursal.Size = new System.Drawing.Size(136, 22);
			this.chkSucursal.TabIndex = 8;
			this.chkSucursal.Text = "Toda Sucursal";
			// 
			// btContabilizar
			// 
			this.btContabilizar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btContabilizar.Location = new System.Drawing.Point(585, 34);
			this.btContabilizar.Name = "btContabilizar";
			this.btContabilizar.Size = new System.Drawing.Size(79, 26);
			this.btContabilizar.TabIndex = 10;
			this.btContabilizar.Text = "&Contabilizar";
			this.btContabilizar.Click += new System.EventHandler(this.btContabilizar_Click);
			this.btContabilizar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btContabilizar_MouseUp);
			// 
			// btEgreso
			// 
			this.btEgreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btEgreso.Location = new System.Drawing.Point(672, 34);
			this.btEgreso.Name = "btEgreso";
			this.btEgreso.Size = new System.Drawing.Size(66, 26);
			this.btEgreso.TabIndex = 11;
			this.btEgreso.Text = "&Egresos";
			this.btEgreso.Click += new System.EventHandler(this.btEgreso_Click);
			this.btEgreso.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btEgreso_MouseDown);
			// 
			// btEditar
			// 
			this.btEditar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btEditar.Location = new System.Drawing.Point(72, 34);
			this.btEditar.Name = "btEditar";
			this.btEditar.Size = new System.Drawing.Size(56, 25);
			this.btEditar.TabIndex = 12;
			this.btEditar.Text = "Editar";
			this.btEditar.Click += new System.EventHandler(this.btEditar_Click);
			// 
			// cmbRolGrupo
			// 
			this.cmbRolGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRolGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRolGrupo.DataMember = "RolGrupo";
			this.cmbRolGrupo.DataSource = this.cdsRolGrupo;
			this.cmbRolGrupo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "RolTipo";
			ultraGridColumn216.Header.VisiblePosition = 1;
			ultraGridColumn216.Hidden = true;
			ultraGridColumn217.Header.VisiblePosition = 0;
			ultraGridColumn217.Width = 133;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn216,
																										 ultraGridColumn217});
			this.cmbRolGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbRolGrupo.DisplayMember = "Grupo";
			this.cmbRolGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRolGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRolGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRolGrupo.Location = new System.Drawing.Point(672, 9);
			this.cmbRolGrupo.Name = "cmbRolGrupo";
			this.cmbRolGrupo.Size = new System.Drawing.Size(152, 21);
			this.cmbRolGrupo.TabIndex = 15;
			this.cmbRolGrupo.ValueMember = "idRolGrupo";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// chkTodoGrupo
			// 
			this.chkTodoGrupo.Checked = true;
			this.chkTodoGrupo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodoGrupo.Location = new System.Drawing.Point(584, 9);
			this.chkTodoGrupo.Name = "chkTodoGrupo";
			this.chkTodoGrupo.Size = new System.Drawing.Size(88, 21);
			this.chkTodoGrupo.TabIndex = 16;
			this.chkTodoGrupo.Text = "Todo Grupo";
			// 
			// btEnvio
			// 
			this.btEnvio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btEnvio.Location = new System.Drawing.Point(832, 34);
			this.btEnvio.Name = "btEnvio";
			this.btEnvio.Size = new System.Drawing.Size(80, 26);
			this.btEnvio.TabIndex = 17;
			this.btEnvio.Text = "E&nvio";
			this.btEnvio.Click += new System.EventHandler(this.btEnvio_Click);
			// 
			// btExcel
			// 
			this.btExcel.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btExcel.Location = new System.Drawing.Point(744, 34);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(80, 26);
			this.btExcel.TabIndex = 18;
			this.btExcel.Text = "E&xcel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// btRubro
			// 
			this.btRubro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRubro.Location = new System.Drawing.Point(832, 9);
			this.btRubro.Name = "btRubro";
			this.btRubro.Size = new System.Drawing.Size(35, 25);
			this.btRubro.TabIndex = 19;
			this.btRubro.Text = "&R";
			this.toolTip1.SetToolTip(this.btRubro, "Muestra el rubro de la columna seleccionada");
			this.btRubro.Click += new System.EventHandler(this.btRubro_Click);
			// 
			// btPersonal
			// 
			this.btPersonal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPersonal.Location = new System.Drawing.Point(880, 9);
			this.btPersonal.Name = "btPersonal";
			this.btPersonal.Size = new System.Drawing.Size(35, 25);
			this.btPersonal.TabIndex = 20;
			this.btPersonal.Text = "&P";
			this.toolTip1.SetToolTip(this.btPersonal, "Muestra el personal de la fila seleccionada");
			this.btPersonal.Click += new System.EventHandler(this.btPersonal_Click);
			// 
			// RolGenera
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(920, 436);
			this.Controls.Add(this.btPersonal);
			this.Controls.Add(this.btRubro);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btEnvio);
			this.Controls.Add(this.chkTodoGrupo);
			this.Controls.Add(this.cmbRolGrupo);
			this.Controls.Add(this.btEditar);
			this.Controls.Add(this.btEgreso);
			this.Controls.Add(this.btContabilizar);
			this.Controls.Add(this.cmbDepartamento);
			this.Controls.Add(this.chkSucursal);
			this.Controls.Add(this.cmbSucursal);
			this.Controls.Add(this.chkDepartamento);
			this.Controls.Add(this.grdRolPago);
			this.Controls.Add(this.btGuardar);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbFecha);
			this.Name = "RolGenera";
			this.Text = "Generacion de Roles";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RolGenera_Closing);
			this.Load += new System.EventHandler(this.RolGenera_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRolPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRolGrupo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{

		}

		private void MostrarColumnas()
		{
			string stFiltro = "";
			if (!chkTodoGrupo.Checked) stFiltro = "And idRolGrupo = " + cmbRolGrupo.Value.ToString();
			string stColumna = "";
			string stConcepto = "";
			int iPos, iTipo, iOrden;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridColumn dc in grdRolPago.DisplayLayout.Bands[0].Columns)
			{
				if (dc.Key.ToLower().StartsWith("ingreso") || dc.Key.ToLower().StartsWith("egreso")
					|| dc.Key.ToLower().StartsWith("provision") || dc.Key.ToLower().StartsWith("valor"))
					dc.Hidden = true;
			}

			string stLista = "Select IsNull(idRolTipo, 0), IsNull(Posicion, 0), IsNull(Orden, 0), " +
				"IsNUll(Concepto, ''), IsNull(Mostrar, 0), IsNull(Modificar, 0) From RolConcepto "
				+ "Where IsNull(Mostrar, 1) = 1 " + stFiltro;

			SqlDataReader dr = Funcion.rEscalarSQL(cdsRolGrupo, stLista, true);
			while(dr.Read())
				#region Lee
			{
				iPos = 0;
				iTipo = dr.GetInt32(0);
				iOrden = 0;
				switch(iTipo)
				{
					case 1:
						stColumna = "Ingreso";
						break;
					case 2:
						stColumna = "Egreso";
						break;
					case 3:
						stColumna = "Provision";
						break;
					case 4:
						stColumna = "Valor";
						break;
				}
			
				iPos = dr.GetInt32(1);
				iOrden = dr.GetInt32(2) + 4; // 4 para nombre, mes, anio, falta (5 -1)

				stConcepto = dr.GetString(3);
				bool bMostrar = dr.GetBoolean(4);
				bool bModifica = dr.GetBoolean(5);

				if (iPos > 0 && iTipo > 0 && bMostrar)
				{
					stColumna += iPos.ToString().Trim();
					grdRolPago.DisplayLayout.Bands[0].Columns[stColumna].Hidden = false;
					grdRolPago.DisplayLayout.Bands[0].Columns[stColumna].Header.VisiblePosition = iOrden;
					grdRolPago.DisplayLayout.Bands[0].Columns[stColumna].Header.Caption = stConcepto;
					if (!bModifica)
						grdRolPago.DisplayLayout.Bands[0].Columns[stColumna].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
					else
						grdRolPago.DisplayLayout.Bands[0].Columns[stColumna].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				}
			}
			#endregion Lee

			#region Observacion
			grdRolPago.DisplayLayout.Bands[0].Columns["Observacion"].Hidden = false;
			grdRolPago.DisplayLayout.Bands[0].Columns["Observacion"].Header.VisiblePosition = 999;
			#endregion
			dr.Close();
		}

		private void RolGenera_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Rol Genera'";
			Funcion.EjecutaSQL(cdsRolGrupo, stAudita, true);
			Cursor = Cursors.Default;
//			this.grdConcepto.Width = 0;
			DateTime dtFecha = DateTime.Today;
			cmbFecha.Value = dtFecha;
			string stFiltro = "Mes= " + dtFecha.Month.ToString() + " And Anio = " + dtFecha.Year.ToString();// + " And idGrupo = 1";
			C1.Data.FilterCondition fcPadre = new C1.Data.FilterCondition(this.cdsRolGrupo, "RolPago", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			this.cdsRolGrupo.Fill(fcTotal, false);
			MostrarColumnas();

			#region Filtro de Sucursales
			bool bFiltroSucursal = Funcion.bEjecutaSQL(cdsRolGrupo, "Select IsNull(FiltroSucursal, 0) From SeteoRol");
			if (bFiltroSucursal)
			{
				string iSuc = Funcion.sEscalarSQL(cdsRolGrupo, "Exec VerificaUsuarioRol");
				if (iSuc != "")
				{
					cdvSucursal.RowFilter = "idRolSucursal in(" + iSuc.ToString() + " )";
					chkSucursal.Checked = false;
					chkSucursal.Enabled = false;
					cmbSucursal.Value = iSuc;
				}
			}
			#endregion Filtro de Sucursales
		}

		private bool Verifica()
		{
			if (!chkDepartamento.Checked && cmbDepartamento.Text.Trim().Length == 0)
			{
        MessageBox.Show("Ingrese departamento", "Información");
				return false;
			}
			if (!chkSucursal.Checked && cmbSucursal.Text.Trim().Length == 0)
			{
				MessageBox.Show("Ingrese Sucursal", "Información");
				return false;
			}
			string stSelect = "Select Count(*) From RolConcepto Where Calculo = 1 And "
				+ "(LEN(RTRIM(IsNull(Formula, ''))) = 0 Or LEN(RTRIM(IsNull(FormulaSQL, ''))) = 0)";
			int iErrorFormula = Funcion.iEscalarSQL(cdsRolGrupo, stSelect, true);
			if (iErrorFormula > 0)
			{
				MessageBox.Show("Existe un Rubro que esta con visto en formula y no tiene una formula", "Información");
				stSelect = "Select idRolConcepto From RolConcepto Where Calculo = 1 And "
					+ "(LEN(RTRIM(IsNull(Formula, ''))) = 0 Or LEN(RTRIM(IsNull(FormulaSQL, ''))) = 0)";
				int IdRolConcepto = Funcion.iEscalarSQL(cdsRolGrupo, stSelect, true);
				RolConcepto miConcepto = new RolConcepto(IdRolConcepto);
				miConcepto.MdiParent = this.MdiParent;
				miConcepto.Show();
				return false;
			}

			DateTime dtFecha = (DateTime) this.cmbFecha.Value;
			stSelect = "Select Count(*) From RolPago Where Bloqueado = 1 And Mes = "
				+ dtFecha.Month.ToString() + " And Anio = " + dtFecha.Year.ToString();
			int iCuenta = Funcion.iEscalarSQL(cdsRolGrupo, stSelect, true);
			if (iCuenta >= 1)
			{
				MessageBox.Show("Registros Bloqueados no puede Realizar Modificaciones", "Información");
				return false;
			}
			return true;
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("278", cdsRolGrupo))
			{
				MessageBox.Show("No tiene acceso a la Edicion de Roles",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!VerificaBloqueo()) return;
			errorProvider.SetError(cmbRolGrupo, "");
			if (cmbRolGrupo.Value == null || (int) cmbRolGrupo.Value == 0)
			{
				errorProvider.SetError(cmbRolGrupo, "Ingrese Grupo");
				return;
			}
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) this.cmbFecha.Value;
			string stSelect = "EXEC RolGenera " + dtFecha.Month.ToString()
				+ ", " + dtFecha.Year.ToString() + ", ";

			stSelect += cmbRolGrupo.Value.ToString() + ", ";

			if (!chkDepartamento.Checked)
				stSelect += cmbDepartamento.Value.ToString();
			else
				stSelect += "0";
			if (!chkSucursal.Checked)
				stSelect += ", " + cmbSucursal.Value.ToString();
			else
				stSelect += ", 0";

			Funcion.EjecutaSQL(cdsRolGrupo, stSelect, true);
			Funcion.EjecutaSQL(cdsRolGrupo, stSelect, true);
			Funcion.EjecutaSQL(cdsRolGrupo, stSelect, true);
			Funcion.EjecutaSQL(cdsRolGrupo, stSelect, true);

			string stSelect1 = "EXEC RolDecimos " + dtFecha.Month.ToString()
				+ ", " + dtFecha.Year.ToString() + ", 1";
			Funcion.EjecutaSQL(cdsRolGrupo, stSelect1);
			Funcion.EjecutaSQL(cdsRolGrupo, stSelect1);
			Funcion.EjecutaSQL(cdsRolGrupo, stSelect1);
			Funcion.EjecutaSQL(cdsRolGrupo, stSelect1);

			string stFiltro = "Mes= " + dtFecha.Month.ToString()
				+ " And Anio = " + dtFecha.Year.ToString();// + " And idGrupo = 1";
			if (!chkSucursal.Checked)
				stFiltro += " And idSucursal = " + cmbSucursal.Value.ToString();
			if (!chkDepartamento.Checked)
				stFiltro += " And idDepartamento = " + cmbDepartamento.Value.ToString();
			if (!chkTodoGrupo.Checked)
				stFiltro += " And idRolGrupo = " + cmbRolGrupo.Value.ToString();
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsRolGrupo, "RolPago", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			cdsRolGrupo.Fill(fcTotal, false);
			Cursor = Cursors.Default;
			MessageBox.Show("Rol de Pagos Generado", "Información");
		}

		private void btGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("278", cdsRolGrupo))
			{
				MessageBox.Show("No tiene acceso a la Edicion de Roles",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (!VerificaBloqueo()) return;
			if (!Verifica()) return;
			this.cdsRolGrupo.Update();
		}

		private void cdsRolGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRolGrupo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsRolTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRolTabla.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void btContabilizar_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbRolGrupo, "");
			if (cmbRolGrupo.Value == null)
			{
				errorProvider.SetError(cmbRolGrupo, "Seleccione Grupo");
				return;
			}
			string stGrupo = cmbRolGrupo.Text;
			if (!Funcion.Permiso("279", cdsRolGrupo))
			{
				MessageBox.Show("No tiene acceso a la Contabilizacion",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show(string.Format("¿Desea contabilizar rol de Grupo {0}?", stGrupo)
				, "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stSelect = " " + dtFecha.Month.ToString()
				+ ", " + dtFecha.Year.ToString() + " ";
			if (!chkSucursal.Checked) stSelect += ", " + cmbSucursal.Value.ToString();
			else stSelect += ", 0";

			if (!chkDepartamento.Checked) stSelect += ", " + cmbDepartamento.Value.ToString();
			else stSelect += ", 0";

			stSelect += ", " + cmbRolGrupo.Value.ToString();

			string stTitulo = Funcion.sEscalarSQL(cdsRolGrupo, "Exec RolTitulo" + stSelect, true);
			if (stTitulo.StartsWith("Ya Existe"))
			if (DialogResult.No == MessageBox.Show(stTitulo + "\n¿Desea Sobrescribir rol?"
				, "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			Funcion.EjecutaSQL(cdsRolGrupo, "Exec RolContabiliza" + stSelect, true);
			MessageBox.Show("Rol Contabilizado", "Información");
		}

		private bool VerificaFiltro()
		{
			bool bOk = true;
			errorProvider.SetError(cmbSucursal, "");
			errorProvider.SetError(cmbDepartamento, "");
			errorProvider.SetError(cmbRolGrupo, "");
			if (!chkSucursal.Checked)
				if (cmbSucursal.Value == null)
				{
					errorProvider.SetError(cmbSucursal, "Ingrese Sucursal");
					bOk = false;
				}
			if (!chkDepartamento.Checked)
				if (cmbDepartamento.Value == null)
				{
					errorProvider.SetError(cmbDepartamento, "Ingrese Departamento");
					bOk = false;
				}
			if (cmbRolGrupo.Value == null)
			{
				errorProvider.SetError(cmbRolGrupo, "Ingrese Grupo");
				bOk = false;
			}
			return bOk;
		}

		private void btEgreso_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("279", cdsRolGrupo))
			{
				MessageBox.Show("No tiene acceso a la Contabilizacion",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (!VerificaFiltro()) return;

			if (DialogResult.No == MessageBox.Show("¿Desea Genera Rol Mensual?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			string stSucursal = ", 0";
			string stDepartamento = ", 0";
			string stGrupo = ", " + cmbRolGrupo.Value.ToString();
			if (!chkSucursal.Checked) stSucursal = ", " + cmbSucursal.Value.ToString();
			if (!chkDepartamento.Checked) stDepartamento = ", " + cmbDepartamento.Value.ToString();

			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stSelect = "Exec RolEgreso " + dtFecha.Month.ToString()
				+ ", " + dtFecha.Year.ToString() + stSucursal + stDepartamento + stGrupo;

			Funcion.EjecutaSQL(cdsRolGrupo, stSelect, true);
			MessageBox.Show("Egresos Generados", "Información");
		}

		private bool VerificaBloqueo()
		{
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			int iMes = dtFecha.Month;
			int iAnio = dtFecha.Year;

			string stSelect = "Select Count(*) From RolPago Where Mes = "
				+ iMes.ToString() + " And Anio = " + iAnio.ToString() + " And Bloqueado = 1";
			int iCuenta = Funcion.iEscalarSQL(cdsRolTabla, stSelect, true);
			if (iCuenta > 0)
			{
				MessageBox.Show("No puede Modificar Roles despues de Bloquear", "Información");
				return false;
			}
      return true;
		}

		private void btEditar_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("278", cdsRolGrupo))
			{
				MessageBox.Show("No tiene acceso a la Edicion de Roles",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			DateTime dtFecha = (DateTime) cmbFecha.Value;
			errorProvider.SetError(cmbRolGrupo, "");
			if (cmbRolGrupo.Value == null || (int) cmbRolGrupo.Value == 0)
			{
				errorProvider.SetError(cmbRolGrupo, "Ingrese Grupo");
				return;
			}

			string stFiltro = "Mes= " + dtFecha.Month.ToString()
				+ " And Anio = " + dtFecha.Year.ToString();// + " And idGrupo = 1";

			stFiltro += " And idRolGrupo = " + cmbRolGrupo.Value.ToString();
			if (!chkSucursal.Checked)
				stFiltro += " And idSucursal = " + cmbSucursal.Value.ToString();
			if (!chkDepartamento.Checked)
				stFiltro += " And idDepartamento = " + cmbDepartamento.Value.ToString();
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsRolGrupo, "RolPago", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			cdsRolGrupo.Fill(fcTotal, false);
			MessageBox.Show("Rol de Pagos Mostrado", "Información");
		}

		private void grdRolPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.Alt) 
				return;
			if (grdRolPago.ActiveCell == null) return;
			switch(e.KeyValue)
			{
				case (int) Keys.PageDown:
					if(grdRolPago.ActiveCell.DroppedDown == false)
						grdRolPago.ActiveCell.DroppedDown = true;
					break;

				case (int) Keys.Up:
					if (grdRolPago.ActiveCell.DroppedDown) return;

					grdRolPago.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grdRolPago.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					grdRolPago.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (grdRolPago.ActiveCell.DroppedDown) return;
			
					grdRolPago.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grdRolPago.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					grdRolPago.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Enter:
					grdRolPago.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grdRolPago.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					grdRolPago.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
			}
		}

		private void RolGenera_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Rol Genera'";
			Funcion.EjecutaSQL(cdsRolGrupo, stAudita, true);
		}

		private void btContabilizar_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (!Funcion.Permiso("225", cdsRolGrupo))
			{
				MessageBox.Show("No tiene acceso a la Contabilizacion",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (e.Button != MouseButtons.Right) return;
			MostrarColumnas();
		}

		private void btEgreso_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (!Funcion.Permiso("225", cdsRolGrupo))
			{
				MessageBox.Show("No tiene acceso a la Contabilizacion",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (e.Button != MouseButtons.Right) return;

			if (!VerificaFiltro()) return;

			if (DialogResult.No == MessageBox.Show("¿Desea Genera Rol Quincena?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			string stSucursal = ", 0";
			string stDepartamento = ", 0";
			string stGrupo = ", " + cmbRolGrupo.Value.ToString();
			if (!chkSucursal.Checked) stSucursal = ", " + cmbSucursal.Value.ToString();
			if (!chkDepartamento.Checked) stDepartamento = ", " + cmbDepartamento.Value.ToString();

			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stSelect = "Exec RolEgresoQuincena " + dtFecha.Month.ToString()
				+ ", " + dtFecha.Year.ToString() + stSucursal + stDepartamento + stGrupo;

			Funcion.EjecutaSQL(cdsRolGrupo, stSelect, true);
			MessageBox.Show("Egresos Generados", "Información");
		}

		private void cmbSucursal_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btEnvio_Click(object sender, System.EventArgs e)
		{
//			if (!Funcion.Permiso("279", cdsRolGrupo))
//			{
//				MessageBox.Show("No tiene acceso a la Envio de Informacion",
//					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
//				return;
//			}
			
			if (DialogResult.No == MessageBox.Show("¿Desea Enviar Datos?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			string stSucursal = ", 0";
			string stDepartamento = ", 0";
			if (!chkSucursal.Checked) stSucursal = ", " + cmbSucursal.Value.ToString();
			if (!chkDepartamento.Checked) stDepartamento = ", " + cmbDepartamento.Value.ToString();

			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stSelect = "Exec Agro_EnvioRol " + dtFecha.Month.ToString()
				+ ", " + dtFecha.Year.ToString();

			Funcion.EjecutaSQL(cdsRolGrupo, stSelect, true);
			MessageBox.Show("Rol Enviado", "Información");

		}
		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(grdRolPago, "RolGeneral.xls");
			if (DialogResult.Yes == MessageBox.Show(
			"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "RolGeneral.xls");

		}

		private void btRubro_Click(object sender, System.EventArgs e)
		{
			int idConcepto = 0;
			int idRolGrupo = 0;
			if(!chkTodoGrupo.Checked && cmbRolGrupo.Value != null && (int)cmbRolGrupo.Value > 0) idRolGrupo = (int)cmbRolGrupo.Value;
			if (grdRolPago.ActiveCell == null)
			{
				MessageBox.Show("Seleccione columna que desea ver el Rubro", "Información");
				return;
			}
			string stColumna = grdRolPago.ActiveCell.Column.ToString().ToLower();
			int idTipo = 4;
			string stTipo = "valor";
			if (stColumna.StartsWith("ingreso")) {idTipo = 1; stTipo = "ingreso";}
			else if (stColumna.StartsWith("egreso")) {idTipo = 2; stTipo = "egreso";}
			else if (stColumna.StartsWith("provision")) {idTipo = 3; stTipo = "provision";}
			else if (stColumna.StartsWith("valor")) {idTipo = 4; stTipo = "valor";}
			else
			{
				MessageBox.Show("Columna no contiene un Rubro", "Información");
				return;
			}
			int iPos = stColumna.IndexOf(stTipo);
			stTipo = stColumna.Substring(stTipo.Length);
			iPos = int.Parse(stTipo);
			string stSelect = string.Format("Select Count(*) From RolConcepto Where idRolTipo = {0} And Posicion = {1} And (idRolGrupo = {2} Or {2} = 0)",
				idTipo, iPos, idRolGrupo);

			idConcepto = Funcion.iEscalarSQL(cdsRolGrupo, stSelect);
			if (idConcepto == 0)
			{
				MessageBox.Show("Rubro no existe en grupo seleccionado", "Información");
				return;
			}
			stSelect = stSelect.Replace("Count(*)", "Top 1 idRolConcepto");
			idConcepto = Funcion.iEscalarSQL(cdsRolGrupo, stSelect);
			RolConcepto miConcepto = new RolConcepto(idConcepto);
			miConcepto.MdiParent = MdiParent;
			miConcepto.Show();
		}

		private void btPersonal_Click(object sender, System.EventArgs e)
		{
			if (grdRolPago.ActiveRow == null)
			{
				MessageBox.Show("Seleccione fila que desea ver la informacion de la persona", "Información");
				return;
			}
			int IdPersona = (int) this.grdRolPago.ActiveRow.Cells["idPersonal"].Value;
			RolPersonal miPersona = new RolPersonal(IdPersona);
			miPersona.MdiParent = MdiParent;
			miPersona.Show();
		}



	}
}
