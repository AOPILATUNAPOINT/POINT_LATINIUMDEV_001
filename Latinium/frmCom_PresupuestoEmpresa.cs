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
	/// Descripción breve de frmCom_PresupuestoEmpresa.
	/// </summary>
	public class frmCom_PresupuestoEmpresa : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesGenerarNuevo;
		private System.Windows.Forms.MenuItem mnuAccionesCopiarUltimoMes;
		private System.Windows.Forms.MenuItem mnuAccionesAgregarNuevoLocal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnio;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalleGEx;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Button btnCargosDeVentas;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalleContado;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalleCredito;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCarpa;
		private System.Windows.Forms.ContextMenu contextMenuExportar;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.MenuItem menuItem4;
		private System.Windows.Forms.Button BtnExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_PresupuestoEmpresa()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 1, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 2, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 3, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 4, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 5, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 6, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 7, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 8, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 9, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 10, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 11, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 13, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_PresupuestoEmpresa));
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 1, true);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 2, true);
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 3, true);
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 4, true);
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 5, true);
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 6, true);
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 7, true);
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 8, true);
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 9, true);
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 10, true);
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 11, true);
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 13, true);
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings27 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 1, true);
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings28 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 2, true);
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings29 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 3, true);
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings30 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 4, true);
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings31 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 5, true);
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings32 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 6, true);
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings33 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 7, true);
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings34 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 8, true);
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings35 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 9, true);
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings36 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 10, true);
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings37 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 11, true);
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings38 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings39 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 13, true);
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings40 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 1, true);
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings41 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 2, true);
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings42 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 3, true);
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings43 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 4, true);
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings44 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 5, true);
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings45 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 6, true);
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings46 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 7, true);
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings47 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 8, true);
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings48 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 9, true);
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings49 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 10, true);
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings50 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 11, true);
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings51 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings52 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 13, true);
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings53 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 1, true);
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings54 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 2, true);
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings55 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 3, true);
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings56 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 4, true);
			Infragistics.Win.Appearance appearance155 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings57 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 5, true);
			Infragistics.Win.Appearance appearance156 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings58 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 6, true);
			Infragistics.Win.Appearance appearance157 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings59 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 7, true);
			Infragistics.Win.Appearance appearance158 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings60 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 8, true);
			Infragistics.Win.Appearance appearance159 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings61 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 9, true);
			Infragistics.Win.Appearance appearance160 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings62 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 10, true);
			Infragistics.Win.Appearance appearance161 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings63 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 11, true);
			Infragistics.Win.Appearance appearance162 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings64 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance163 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings65 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 13, true);
			Infragistics.Win.Appearance appearance164 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance165 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance166 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance167 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance168 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance169 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance170 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance171 = new Infragistics.Win.Appearance();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.btnVer = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnComandos = new System.Windows.Forms.Button();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesGenerarNuevo = new System.Windows.Forms.MenuItem();
			this.mnuAccionesCopiarUltimoMes = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAgregarNuevoLocal = new System.Windows.Forms.MenuItem();
			this.txtAnio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.uGridDetalleGEx = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDetalleContado = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnCargosDeVentas = new System.Windows.Forms.Button();
			this.uGridDetalleCredito = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridCarpa = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.BtnExcel = new System.Windows.Forms.Button();
			this.contextMenuExportar = new System.Windows.Forms.ContextMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this.menuItem4 = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleGEx)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleContado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCarpa)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance1;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.Caption = "LOCAL";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 160;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.CellAppearance = appearance2;
			ultraGridColumn2.Format = "#,##0.00";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 75;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance3;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 75;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 75;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance5;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 75;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance6;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 75;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance7;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 75;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance8;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 75;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance9;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 75;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance10;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 77;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance11;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 75;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance12;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 75;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance13;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 75;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance14;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 80;
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
																										 ultraGridColumn14});
			ultraGridBand1.Header.Caption = "General";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance15;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance16;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance17;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance18;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance19;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance20;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance21;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance22;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance23;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance24;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance25;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance26;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance27;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
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
																																															summarySettings13});
			ultraGridBand1.SummaryFooterCaption = "TOTALES";
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance28;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance29.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance29;
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance30;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance31.FontData.BoldAsString = "True";
			appearance31.FontData.Name = "Arial";
			appearance31.FontData.SizeInPoints = 8F;
			appearance31.ForeColor = System.Drawing.Color.White;
			appearance31.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance31;
			appearance32.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance32.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance34;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(7, 48);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(1163, 160);
			this.uGridDetalle.TabIndex = 16;
			this.uGridDetalle.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDetalle_BeforeCellUpdate);
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(System.Decimal);
			ultraDataColumn2.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn3.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn4.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn5.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn6.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn13.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
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
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel4.Location = new System.Drawing.Point(8, 12);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(23, 15);
			this.ultraLabel4.TabIndex = 355;
			this.ultraLabel4.Text = "Año";
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(160, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 659;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1170, 8);
			this.groupBox1.TabIndex = 661;
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
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(256, 9);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 717;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(1007, 8);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.TabIndex = 726;
			this.btnComandos.Text = "Acciones";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click);
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesGenerarNuevo,
																																								this.mnuAccionesCopiarUltimoMes,
																																								this.mnuAccionesAgregarNuevoLocal});
			this.mnuAcciones.Popup += new System.EventHandler(this.mnuAcciones_Popup);
			// 
			// mnuAccionesGenerarNuevo
			// 
			this.mnuAccionesGenerarNuevo.Enabled = false;
			this.mnuAccionesGenerarNuevo.Index = 0;
			this.mnuAccionesGenerarNuevo.Text = "Nuevo Presupuesto Anual";
			this.mnuAccionesGenerarNuevo.Click += new System.EventHandler(this.mnuAccionesGenerarNuevo_Click);
			// 
			// mnuAccionesCopiarUltimoMes
			// 
			this.mnuAccionesCopiarUltimoMes.Index = 1;
			this.mnuAccionesCopiarUltimoMes.Text = "Copiar Ultimo Mes";
			this.mnuAccionesCopiarUltimoMes.Click += new System.EventHandler(this.mnuAccionesCopiarUltimoMes_Click);
			// 
			// mnuAccionesAgregarNuevoLocal
			// 
			this.mnuAccionesAgregarNuevoLocal.Index = 2;
			this.mnuAccionesAgregarNuevoLocal.Text = "Agregar Nuevo Local";
			this.mnuAccionesAgregarNuevoLocal.Click += new System.EventHandler(this.mnuAccionesAgregarNuevoLocal_Click);
			// 
			// txtAnio
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAnio.Appearance = appearance35;
			this.txtAnio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAnio.FormatString = "";
			this.txtAnio.Location = new System.Drawing.Point(48, 8);
			this.txtAnio.MaskInput = "nnnn";
			this.txtAnio.Name = "txtAnio";
			this.txtAnio.PromptChar = ' ';
			this.txtAnio.Size = new System.Drawing.Size(96, 22);
			this.txtAnio.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAnio.TabIndex = 734;
			this.txtAnio.Click += new System.EventHandler(this.txtAnio_Click);
			this.txtAnio.Enter += new System.EventHandler(this.txtAnio_Enter);
			// 
			// uGridDetalleGEx
			// 
			this.uGridDetalleGEx.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalleGEx.DataSource = this.ultraDataSource2;
			appearance36.BackColor = System.Drawing.Color.White;
			this.uGridDetalleGEx.DisplayLayout.Appearance = appearance36;
			this.uGridDetalleGEx.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.Caption = "LOCAL";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Width = 166;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance37;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 74;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance38;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Width = 74;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance39;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Width = 74;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance40;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Width = 74;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance41;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridColumn20.Width = 74;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance42;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Width = 74;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance43;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 7;
			ultraGridColumn22.Width = 74;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance44;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 8;
			ultraGridColumn23.Width = 76;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance45;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 9;
			ultraGridColumn24.Width = 78;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance46;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 10;
			ultraGridColumn25.Width = 76;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance47;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 11;
			ultraGridColumn26.Width = 76;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance48;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 12;
			ultraGridColumn27.Width = 76;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance49;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 13;
			ultraGridColumn28.Width = 76;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn28});
			ultraGridBand2.Header.Caption = "Garantia Extendida";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance50;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance51;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance52;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance53;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance54;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance55;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance56;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance57;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance58;
			summarySettings22.DisplayFormat = "{0: #,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance59;
			summarySettings23.DisplayFormat = "{0: #,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance60;
			summarySettings24.DisplayFormat = "{0: #,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance61;
			summarySettings25.DisplayFormat = "{0: #,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance62;
			summarySettings26.DisplayFormat = "{0: #,##0.00}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings26});
			ultraGridBand2.SummaryFooterCaption = "TOTALES";
			this.uGridDetalleGEx.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance63.ForeColor = System.Drawing.Color.Black;
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleGEx.DisplayLayout.Override.ActiveRowAppearance = appearance63;
			this.uGridDetalleGEx.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalleGEx.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDetalleGEx.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance64.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalleGEx.DisplayLayout.Override.CardAreaAppearance = appearance64;
			appearance65.ForeColor = System.Drawing.Color.Black;
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleGEx.DisplayLayout.Override.CellAppearance = appearance65;
			this.uGridDetalleGEx.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance66.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance66.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance66.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance66.FontData.BoldAsString = "True";
			appearance66.FontData.Name = "Arial";
			appearance66.FontData.SizeInPoints = 8F;
			appearance66.ForeColor = System.Drawing.Color.White;
			appearance66.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalleGEx.DisplayLayout.Override.HeaderAppearance = appearance66;
			appearance67.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance67.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance67.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleGEx.DisplayLayout.Override.RowAlternateAppearance = appearance67;
			appearance68.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance68.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance68.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleGEx.DisplayLayout.Override.RowSelectorAppearance = appearance68;
			appearance69.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance69.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance69.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleGEx.DisplayLayout.Override.SelectedRowAppearance = appearance69;
			this.uGridDetalleGEx.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalleGEx.Location = new System.Drawing.Point(7, 368);
			this.uGridDetalleGEx.Name = "uGridDetalleGEx";
			this.uGridDetalleGEx.Size = new System.Drawing.Size(1163, 160);
			this.uGridDetalleGEx.TabIndex = 735;
			this.uGridDetalleGEx.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDetalleGEx_BeforeCellUpdate);
			this.uGridDetalleGEx.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalleGEx_KeyDown);
			this.uGridDetalleGEx.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalleGEx_AfterCellUpdate);
			this.uGridDetalleGEx.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalleGEx_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn17.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn18.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn19.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn28});
			// 
			// ultraDataSource3
			// 
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
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn42});
			// 
			// uGridDetalleContado
			// 
			this.uGridDetalleContado.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalleContado.DataSource = this.ultraDataSource3;
			appearance70.BackColor = System.Drawing.Color.White;
			this.uGridDetalleContado.DisplayLayout.Appearance = appearance70;
			this.uGridDetalleContado.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn29.Header.Caption = "LOCAL";
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Width = 166;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance71;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 74;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance72;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Width = 74;
			appearance73.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance73;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.VisiblePosition = 3;
			ultraGridColumn32.Width = 74;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance74;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 4;
			ultraGridColumn33.Width = 74;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance75;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 5;
			ultraGridColumn34.Width = 74;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance76;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 6;
			ultraGridColumn35.Width = 74;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance77;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.VisiblePosition = 7;
			ultraGridColumn36.Width = 74;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance78;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 8;
			ultraGridColumn37.Width = 76;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance79;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.VisiblePosition = 9;
			ultraGridColumn38.Width = 78;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance80;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 10;
			ultraGridColumn39.Width = 76;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance81;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.VisiblePosition = 11;
			ultraGridColumn40.Width = 76;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance82;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 12;
			ultraGridColumn41.Width = 76;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance83;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.VisiblePosition = 13;
			ultraGridColumn42.Width = 76;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn42});
			ultraGridBand3.Header.Caption = "Contado";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings27.Appearance = appearance84;
			summarySettings27.DisplayFormat = "{0: #,##0.00}";
			summarySettings27.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings28.Appearance = appearance85;
			summarySettings28.DisplayFormat = "{0: #,##0.00}";
			summarySettings28.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance86.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings29.Appearance = appearance86;
			summarySettings29.DisplayFormat = "{0: #,##0.00}";
			summarySettings29.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance87.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings30.Appearance = appearance87;
			summarySettings30.DisplayFormat = "{0: #,##0.00}";
			summarySettings30.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance88.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings31.Appearance = appearance88;
			summarySettings31.DisplayFormat = "{0: #,##0.00}";
			summarySettings31.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance89.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings32.Appearance = appearance89;
			summarySettings32.DisplayFormat = "{0: #,##0.00}";
			summarySettings32.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance90.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings33.Appearance = appearance90;
			summarySettings33.DisplayFormat = "{0: #,##0.00}";
			summarySettings33.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance91.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings34.Appearance = appearance91;
			summarySettings34.DisplayFormat = "{0: #,##0.00}";
			summarySettings34.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance92.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings35.Appearance = appearance92;
			summarySettings35.DisplayFormat = "{0: #,##0.00}";
			summarySettings35.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance93.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings36.Appearance = appearance93;
			summarySettings36.DisplayFormat = "{0: #,##0.00}";
			summarySettings36.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings37.Appearance = appearance94;
			summarySettings37.DisplayFormat = "{0: #,##0.00}";
			summarySettings37.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings38.Appearance = appearance95;
			summarySettings38.DisplayFormat = "{0: #,##0.00}";
			summarySettings38.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance96.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings39.Appearance = appearance96;
			summarySettings39.DisplayFormat = "{0: #,##0.00}";
			summarySettings39.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings39});
			ultraGridBand3.SummaryFooterCaption = "TOTALES";
			this.uGridDetalleContado.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance97.ForeColor = System.Drawing.Color.Black;
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleContado.DisplayLayout.Override.ActiveRowAppearance = appearance97;
			this.uGridDetalleContado.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalleContado.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDetalleContado.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance98.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalleContado.DisplayLayout.Override.CardAreaAppearance = appearance98;
			appearance99.ForeColor = System.Drawing.Color.Black;
			appearance99.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleContado.DisplayLayout.Override.CellAppearance = appearance99;
			this.uGridDetalleContado.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance100.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance100.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance100.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance100.FontData.BoldAsString = "True";
			appearance100.FontData.Name = "Arial";
			appearance100.FontData.SizeInPoints = 8F;
			appearance100.ForeColor = System.Drawing.Color.White;
			appearance100.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalleContado.DisplayLayout.Override.HeaderAppearance = appearance100;
			appearance101.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance101.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance101.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleContado.DisplayLayout.Override.RowAlternateAppearance = appearance101;
			appearance102.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance102.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance102.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleContado.DisplayLayout.Override.RowSelectorAppearance = appearance102;
			appearance103.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance103.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance103.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleContado.DisplayLayout.Override.SelectedRowAppearance = appearance103;
			this.uGridDetalleContado.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalleContado.Location = new System.Drawing.Point(7, 528);
			this.uGridDetalleContado.Name = "uGridDetalleContado";
			this.uGridDetalleContado.Size = new System.Drawing.Size(1163, 160);
			this.uGridDetalleContado.TabIndex = 736;
			this.uGridDetalleContado.Visible = false;
			this.uGridDetalleContado.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDetalleContado_BeforeCellUpdate);
			this.uGridDetalleContado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalleContado_KeyDown);
			this.uGridDetalleContado.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalleContado_AfterCellUpdate);
			this.uGridDetalleContado.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalleContado_InitializeLayout);
			// 
			// btnCargosDeVentas
			// 
			this.btnCargosDeVentas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCargosDeVentas.Image = ((System.Drawing.Image)(resources.GetObject("btnCargosDeVentas.Image")));
			this.btnCargosDeVentas.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCargosDeVentas.Location = new System.Drawing.Point(874, 8);
			this.btnCargosDeVentas.Name = "btnCargosDeVentas";
			this.btnCargosDeVentas.Size = new System.Drawing.Size(128, 23);
			this.btnCargosDeVentas.TabIndex = 737;
			this.btnCargosDeVentas.Text = "Cargos De Ventas";
			this.btnCargosDeVentas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCargosDeVentas.Click += new System.EventHandler(this.btnCargosDeVentas_Click);
			// 
			// uGridDetalleCredito
			// 
			this.uGridDetalleCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalleCredito.DataSource = this.ultraDataSource3;
			appearance104.BackColor = System.Drawing.Color.White;
			this.uGridDetalleCredito.DisplayLayout.Appearance = appearance104;
			this.uGridDetalleCredito.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn43.Header.Caption = "LOCAL";
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Width = 166;
			appearance105.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance105;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Width = 74;
			appearance106.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance106;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.VisiblePosition = 2;
			ultraGridColumn45.Width = 74;
			appearance107.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance107;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 3;
			ultraGridColumn46.Width = 74;
			appearance108.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance108;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.VisiblePosition = 4;
			ultraGridColumn47.Width = 74;
			appearance109.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance109;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.VisiblePosition = 5;
			ultraGridColumn48.Width = 74;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance110;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.VisiblePosition = 6;
			ultraGridColumn49.Width = 74;
			appearance111.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance111;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.VisiblePosition = 7;
			ultraGridColumn50.Width = 74;
			appearance112.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance112;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.VisiblePosition = 8;
			ultraGridColumn51.Width = 76;
			appearance113.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance113;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.VisiblePosition = 9;
			ultraGridColumn52.Width = 78;
			appearance114.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance114;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.VisiblePosition = 10;
			ultraGridColumn53.Width = 76;
			appearance115.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance115;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.VisiblePosition = 11;
			ultraGridColumn54.Width = 76;
			appearance116.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance116;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.VisiblePosition = 12;
			ultraGridColumn55.Width = 76;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance117.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance117;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.VisiblePosition = 13;
			ultraGridColumn56.Width = 76;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn56});
			ultraGridBand4.Header.Caption = "Crédito";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance118.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings40.Appearance = appearance118;
			summarySettings40.DisplayFormat = "{0: #,##0.00}";
			summarySettings40.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance119.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings41.Appearance = appearance119;
			summarySettings41.DisplayFormat = "{0: #,##0.00}";
			summarySettings41.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance120.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings42.Appearance = appearance120;
			summarySettings42.DisplayFormat = "{0: #,##0.00}";
			summarySettings42.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance121.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings43.Appearance = appearance121;
			summarySettings43.DisplayFormat = "{0: #,##0.00}";
			summarySettings43.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance122.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings44.Appearance = appearance122;
			summarySettings44.DisplayFormat = "{0: #,##0.00}";
			summarySettings44.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance123.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings45.Appearance = appearance123;
			summarySettings45.DisplayFormat = "{0: #,##0.00}";
			summarySettings45.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance124.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings46.Appearance = appearance124;
			summarySettings46.DisplayFormat = "{0: #,##0.00}";
			summarySettings46.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance125.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings47.Appearance = appearance125;
			summarySettings47.DisplayFormat = "{0: #,##0.00}";
			summarySettings47.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance126.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings48.Appearance = appearance126;
			summarySettings48.DisplayFormat = "{0: #,##0.00}";
			summarySettings48.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance127.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings49.Appearance = appearance127;
			summarySettings49.DisplayFormat = "{0: #,##0.00}";
			summarySettings49.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance128.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings50.Appearance = appearance128;
			summarySettings50.DisplayFormat = "{0: #,##0.00}";
			summarySettings50.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance129.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings51.Appearance = appearance129;
			summarySettings51.DisplayFormat = "{0: #,##0.00}";
			summarySettings51.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance130.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings52.Appearance = appearance130;
			summarySettings52.DisplayFormat = "{0: #,##0.00}";
			summarySettings52.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings52});
			ultraGridBand4.SummaryFooterCaption = "TOTALES";
			this.uGridDetalleCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance131.ForeColor = System.Drawing.Color.Black;
			appearance131.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleCredito.DisplayLayout.Override.ActiveRowAppearance = appearance131;
			this.uGridDetalleCredito.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalleCredito.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDetalleCredito.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance132.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalleCredito.DisplayLayout.Override.CardAreaAppearance = appearance132;
			appearance133.ForeColor = System.Drawing.Color.Black;
			appearance133.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleCredito.DisplayLayout.Override.CellAppearance = appearance133;
			this.uGridDetalleCredito.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance134.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance134.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance134.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance134.FontData.BoldAsString = "True";
			appearance134.FontData.Name = "Arial";
			appearance134.FontData.SizeInPoints = 8F;
			appearance134.ForeColor = System.Drawing.Color.White;
			appearance134.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalleCredito.DisplayLayout.Override.HeaderAppearance = appearance134;
			appearance135.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance135.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance135.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCredito.DisplayLayout.Override.RowAlternateAppearance = appearance135;
			appearance136.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance136.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance136.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCredito.DisplayLayout.Override.RowSelectorAppearance = appearance136;
			appearance137.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance137.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance137.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCredito.DisplayLayout.Override.SelectedRowAppearance = appearance137;
			this.uGridDetalleCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalleCredito.Location = new System.Drawing.Point(7, 208);
			this.uGridDetalleCredito.Name = "uGridDetalleCredito";
			this.uGridDetalleCredito.Size = new System.Drawing.Size(1163, 160);
			this.uGridDetalleCredito.TabIndex = 738;
			this.uGridDetalleCredito.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDetalleCredito_BeforeCellUpdate);
			this.uGridDetalleCredito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalleCredito_KeyDown);
			this.uGridDetalleCredito.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalleCredito_AfterCellUpdate);
			this.uGridDetalleCredito.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalleCredito_InitializeLayout);
			// 
			// uGridCarpa
			// 
			this.uGridCarpa.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCarpa.DataSource = this.ultraDataSource3;
			appearance138.BackColor = System.Drawing.Color.White;
			this.uGridCarpa.DisplayLayout.Appearance = appearance138;
			this.uGridCarpa.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn57.Header.Caption = "LOCAL";
			ultraGridColumn57.Header.VisiblePosition = 0;
			ultraGridColumn57.Width = 166;
			appearance139.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance139;
			ultraGridColumn58.Format = "#,##0";
			ultraGridColumn58.Header.VisiblePosition = 1;
			ultraGridColumn58.Width = 74;
			appearance140.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance140;
			ultraGridColumn59.Format = "#,##0";
			ultraGridColumn59.Header.VisiblePosition = 2;
			ultraGridColumn59.Width = 74;
			appearance141.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance141;
			ultraGridColumn60.Format = "#,##0";
			ultraGridColumn60.Header.VisiblePosition = 3;
			ultraGridColumn60.Width = 74;
			appearance142.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance142;
			ultraGridColumn61.Format = "#,##0";
			ultraGridColumn61.Header.VisiblePosition = 4;
			ultraGridColumn61.Width = 74;
			appearance143.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance143;
			ultraGridColumn62.Format = "#,##0";
			ultraGridColumn62.Header.VisiblePosition = 5;
			ultraGridColumn62.Width = 74;
			appearance144.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance144;
			ultraGridColumn63.Format = "#,##0";
			ultraGridColumn63.Header.VisiblePosition = 6;
			ultraGridColumn63.Width = 74;
			appearance145.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance145;
			ultraGridColumn64.Format = "#,##0";
			ultraGridColumn64.Header.VisiblePosition = 7;
			ultraGridColumn64.Width = 74;
			appearance146.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance146;
			ultraGridColumn65.Format = "#,##0";
			ultraGridColumn65.Header.VisiblePosition = 8;
			ultraGridColumn65.Width = 76;
			appearance147.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance147;
			ultraGridColumn66.Format = "#,##0";
			ultraGridColumn66.Header.VisiblePosition = 9;
			ultraGridColumn66.Width = 78;
			appearance148.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance148;
			ultraGridColumn67.Format = "#,##0";
			ultraGridColumn67.Header.VisiblePosition = 10;
			ultraGridColumn67.Width = 76;
			appearance149.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance149;
			ultraGridColumn68.Format = "#,##0";
			ultraGridColumn68.Header.VisiblePosition = 11;
			ultraGridColumn68.Width = 76;
			appearance150.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance150;
			ultraGridColumn69.Format = "#,##0";
			ultraGridColumn69.Header.VisiblePosition = 12;
			ultraGridColumn69.Width = 76;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance151.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance151;
			ultraGridColumn70.Format = "#,##0";
			ultraGridColumn70.Header.VisiblePosition = 13;
			ultraGridColumn70.Width = 76;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn70});
			ultraGridBand5.Header.Caption = "Carpa";
			ultraGridBand5.HeaderVisible = true;
			ultraGridBand5.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance152.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings53.Appearance = appearance152;
			summarySettings53.DisplayFormat = "{0: #,##0}";
			summarySettings53.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance153.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings54.Appearance = appearance153;
			summarySettings54.DisplayFormat = "{0: #,##0}";
			summarySettings54.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance154.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings55.Appearance = appearance154;
			summarySettings55.DisplayFormat = "{0: #,##0}";
			summarySettings55.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance155.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings56.Appearance = appearance155;
			summarySettings56.DisplayFormat = "{0: #,##0}";
			summarySettings56.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance156.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings57.Appearance = appearance156;
			summarySettings57.DisplayFormat = "{0: #,##0}";
			summarySettings57.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance157.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings58.Appearance = appearance157;
			summarySettings58.DisplayFormat = "{0: #,##0}";
			summarySettings58.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance158.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings59.Appearance = appearance158;
			summarySettings59.DisplayFormat = "{0: #,##0}";
			summarySettings59.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance159.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings60.Appearance = appearance159;
			summarySettings60.DisplayFormat = "{0: #,##0}";
			summarySettings60.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance160.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings61.Appearance = appearance160;
			summarySettings61.DisplayFormat = "{0: #,##0}";
			summarySettings61.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance161.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings62.Appearance = appearance161;
			summarySettings62.DisplayFormat = "{0: #,##0}";
			summarySettings62.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance162.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings63.Appearance = appearance162;
			summarySettings63.DisplayFormat = "{0: #,##0}";
			summarySettings63.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance163.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings64.Appearance = appearance163;
			summarySettings64.DisplayFormat = "{0: #,##0}";
			summarySettings64.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance164.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings65.Appearance = appearance164;
			summarySettings65.DisplayFormat = "{0: #,##0}";
			summarySettings65.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings65});
			ultraGridBand5.SummaryFooterCaption = "TOTALES";
			this.uGridCarpa.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance165.ForeColor = System.Drawing.Color.Black;
			appearance165.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCarpa.DisplayLayout.Override.ActiveRowAppearance = appearance165;
			this.uGridCarpa.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCarpa.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCarpa.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance166.BackColor = System.Drawing.Color.Transparent;
			this.uGridCarpa.DisplayLayout.Override.CardAreaAppearance = appearance166;
			appearance167.ForeColor = System.Drawing.Color.Black;
			appearance167.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCarpa.DisplayLayout.Override.CellAppearance = appearance167;
			this.uGridCarpa.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance168.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance168.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance168.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance168.FontData.BoldAsString = "True";
			appearance168.FontData.Name = "Arial";
			appearance168.FontData.SizeInPoints = 8F;
			appearance168.ForeColor = System.Drawing.Color.White;
			appearance168.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCarpa.DisplayLayout.Override.HeaderAppearance = appearance168;
			appearance169.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance169.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance169.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCarpa.DisplayLayout.Override.RowAlternateAppearance = appearance169;
			appearance170.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance170.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance170.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCarpa.DisplayLayout.Override.RowSelectorAppearance = appearance170;
			appearance171.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance171.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance171.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCarpa.DisplayLayout.Override.SelectedRowAppearance = appearance171;
			this.uGridCarpa.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCarpa.Location = new System.Drawing.Point(7, 528);
			this.uGridCarpa.Name = "uGridCarpa";
			this.uGridCarpa.Size = new System.Drawing.Size(1163, 160);
			this.uGridCarpa.TabIndex = 739;
			this.uGridCarpa.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridCarpa_BeforeCellUpdate);
			this.uGridCarpa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridCarpa_KeyDown);
			this.uGridCarpa.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCarpa_AfterCellUpdate);
			this.uGridCarpa.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridCarpa_InitializeLayout);
			// 
			// BtnExcel
			// 
			this.BtnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.BtnExcel.Image = ((System.Drawing.Image)(resources.GetObject("BtnExcel.Image")));
			this.BtnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.BtnExcel.Location = new System.Drawing.Point(1090, 8);
			this.BtnExcel.Name = "BtnExcel";
			this.BtnExcel.TabIndex = 740;
			this.BtnExcel.Text = "Exportar";
			this.BtnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.BtnExcel.Click += new System.EventHandler(this.BtnExcel_Click);
			// 
			// contextMenuExportar
			// 
			this.contextMenuExportar.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																												this.menuItem1,
																																												this.menuItem2,
																																												this.menuItem3,
																																												this.menuItem4});
			this.contextMenuExportar.Popup += new System.EventHandler(this.contextMenuExportar_Popup);
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.Text = "General";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 1;
			this.menuItem2.Text = "Crédito";
			this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 2;
			this.menuItem3.Text = "Garntía Extendida";
			this.menuItem3.Click += new System.EventHandler(this.menuItem3_Click);
			// 
			// menuItem4
			// 
			this.menuItem4.Index = 3;
			this.menuItem4.Text = "Carpa";
			this.menuItem4.Click += new System.EventHandler(this.menuItem4_Click);
			// 
			// frmCom_PresupuestoEmpresa
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(1159, 595);
			this.Controls.Add(this.BtnExcel);
			this.Controls.Add(this.uGridCarpa);
			this.Controls.Add(this.uGridDetalleCredito);
			this.Controls.Add(this.btnCargosDeVentas);
			this.Controls.Add(this.txtAnio);
			this.Controls.Add(this.btnComandos);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.uGridDetalleGEx);
			this.Controls.Add(this.uGridDetalleContado);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmCom_PresupuestoEmpresa";
			this.ShowInTaskbar = false;
			this.Text = "Presupuesto de la Empresa";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCom_PresupuestoEmpresa_KeyDown);
			this.Load += new System.EventHandler(this.frmCom_PresupuestoEmpresa_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleGEx)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleContado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCarpa)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCom_PresupuestoEmpresa_Load(object sender, System.EventArgs e)
		{
			this.txtAnio.MinValue = 2016;
			this.txtAnio.MaxValue = DateTime.Today.Year + 1;
			this.txtAnio.Value = DateTime.Today.Year;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio(this.txtAnio, "Ingrese el periodo", 2016, DateTime.Today.Year + 1)) return;

			string sSQL = string.Format("Exec Com_ConsultaPresupuestoEmpresa {0}", Convert.ToInt32(this.txtAnio.Value));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDetalle);

			sSQL = string.Format("Exec Com_ConsultaPresupuestoEmpresaCredito {0}", Convert.ToInt32(this.txtAnio.Value));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDetalleCredito);

			sSQL = string.Format("Exec Com_ConsultaPresupuestoEmpresaGE {0}", Convert.ToInt32(this.txtAnio.Value));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDetalleGEx);

			sSQL = string.Format("Exec Com_ConsultaPresupuestoEmpresaContado {0}", Convert.ToInt32(this.txtAnio.Value));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDetalleContado);
			
			sSQL = string.Format("Exec Com_ConsultaPresupuestoCarpa {0}", Convert.ToInt32(this.txtAnio.Value));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridCarpa);

			this.lblContador.Text = this.uGridDetalle.Rows.Count + " LOCALES ACTIVOS";

			SumatoriasAlConsultar();
		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "ENERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FEBRERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MARZO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ABRIL", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MAYO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JUNIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JULIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AGOSTO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEPTIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "OCTUBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "NOVIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DICIEMBRE", 2);			
		}

		private void ActualizaValores (string sLocal, string sMes, decimal dValor)
		{
			string sSQL = string.Format("Exec Com_ActualizaValoresPresupuestoEmpresa {0}, '{1}', '{2}', {3}", 
				Convert.ToInt32(this.txtAnio.Value), sLocal, sMes, dValor);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void ActualizaValoresCredito (string sLocal, string sMes, decimal dValor)
		{
			string sSQL = string.Format("Exec Com_ActualizaValoresPresupuestoEmpresaCredito {0}, '{1}', '{2}', {3}", 
				Convert.ToInt32(this.txtAnio.Value), sLocal, sMes, dValor);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void ActualizaValoresGEx (string sLocal, string sMes, decimal dValor)
		{
			string sSQL = string.Format("Exec Com_ActualizaValoresPresupuestoEmpresaGE {0}, '{1}', '{2}', {3}", 
				Convert.ToInt32(this.txtAnio.Value), sLocal, sMes, dValor);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}
		
		private void ActualizaValoresContado (string sLocal, string sMes, decimal dValor)
		{
			string sSQL = string.Format("Exec Com_ActualizaValoresPresupuestoEmpresaContado {0}, '{1}', '{2}', {3}", 
				Convert.ToInt32(this.txtAnio.Value), sLocal, sMes, dValor);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void ActualizaValoresCarpa (string sLocal, string sMes, decimal dValor)
		{
			string sSQL = string.Format("Exec Com_ActualizaValoresPresupuestoEmpresaCarpa {0}, '{1}', '{2}', {3}", 
				Convert.ToInt32(this.txtAnio.Value), sLocal, sMes, dValor);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void SumatoriasAlConsultar()
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
			{
				dr.Cells["TOTAL"].Value = 
					Convert.ToDecimal(dr.Cells["ENERO"].Value) + 
					Convert.ToDecimal(dr.Cells["FEBRERO"].Value) +
					Convert.ToDecimal(dr.Cells["MARZO"].Value) +
					Convert.ToDecimal(dr.Cells["ABRIL"].Value) + 
					Convert.ToDecimal(dr.Cells["MAYO"].Value) +
					Convert.ToDecimal(dr.Cells["JUNIO"].Value) +
					Convert.ToDecimal(dr.Cells["JULIO"].Value) + 
					Convert.ToDecimal(dr.Cells["AGOSTO"].Value) +
					Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value) +
					Convert.ToDecimal(dr.Cells["OCTUBRE"].Value) + 
					Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value) +
					Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalleGEx.Rows.All)
			{
				dr.Cells["TOTAL"].Value = 
					Convert.ToDecimal(dr.Cells["ENERO"].Value) + 
					Convert.ToDecimal(dr.Cells["FEBRERO"].Value) +
					Convert.ToDecimal(dr.Cells["MARZO"].Value) +
					Convert.ToDecimal(dr.Cells["ABRIL"].Value) + 
					Convert.ToDecimal(dr.Cells["MAYO"].Value) +
					Convert.ToDecimal(dr.Cells["JUNIO"].Value) +
					Convert.ToDecimal(dr.Cells["JULIO"].Value) + 
					Convert.ToDecimal(dr.Cells["AGOSTO"].Value) +
					Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value) +
					Convert.ToDecimal(dr.Cells["OCTUBRE"].Value) + 
					Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value) +
					Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalleCredito.Rows.All)
			{
				dr.Cells["TOTAL"].Value = 
					Convert.ToDecimal(dr.Cells["ENERO"].Value) + 
					Convert.ToDecimal(dr.Cells["FEBRERO"].Value) +
					Convert.ToDecimal(dr.Cells["MARZO"].Value) +
					Convert.ToDecimal(dr.Cells["ABRIL"].Value) + 
					Convert.ToDecimal(dr.Cells["MAYO"].Value) +
					Convert.ToDecimal(dr.Cells["JUNIO"].Value) +
					Convert.ToDecimal(dr.Cells["JULIO"].Value) + 
					Convert.ToDecimal(dr.Cells["AGOSTO"].Value) +
					Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value) +
					Convert.ToDecimal(dr.Cells["OCTUBRE"].Value) + 
					Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value) +
					Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalleContado.Rows.All)
			{
				dr.Cells["TOTAL"].Value = 
					Convert.ToDecimal(dr.Cells["ENERO"].Value) + 
					Convert.ToDecimal(dr.Cells["FEBRERO"].Value) +
					Convert.ToDecimal(dr.Cells["MARZO"].Value) +
					Convert.ToDecimal(dr.Cells["ABRIL"].Value) + 
					Convert.ToDecimal(dr.Cells["MAYO"].Value) +
					Convert.ToDecimal(dr.Cells["JUNIO"].Value) +
					Convert.ToDecimal(dr.Cells["JULIO"].Value) + 
					Convert.ToDecimal(dr.Cells["AGOSTO"].Value) +
					Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value) +
					Convert.ToDecimal(dr.Cells["OCTUBRE"].Value) + 
					Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value) +
					Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
			}
		}

		private void Sumatorias(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			e.Cell.Row.Cells["TOTAL"].Value = 
					Convert.ToDecimal(e.Cell.Row.Cells["ENERO"].Value) + 
					Convert.ToDecimal(e.Cell.Row.Cells["FEBRERO"].Value) +
					Convert.ToDecimal(e.Cell.Row.Cells["MARZO"].Value) +
					Convert.ToDecimal(e.Cell.Row.Cells["ABRIL"].Value) + 
					Convert.ToDecimal(e.Cell.Row.Cells["MAYO"].Value) +
					Convert.ToDecimal(e.Cell.Row.Cells["JUNIO"].Value) +
					Convert.ToDecimal(e.Cell.Row.Cells["JULIO"].Value) + 
					Convert.ToDecimal(e.Cell.Row.Cells["AGOSTO"].Value) +
					Convert.ToDecimal(e.Cell.Row.Cells["SEPTIEMBRE"].Value) +
					Convert.ToDecimal(e.Cell.Row.Cells["OCTUBRE"].Value) + 
					Convert.ToDecimal(e.Cell.Row.Cells["NOVIEMBRE"].Value) +
					Convert.ToDecimal(e.Cell.Row.Cells["DICIEMBRE"].Value);			
		}

		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ENERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "FEBRERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MARZO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ABRIL", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MAYO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JUNIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JULIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "AGOSTO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "SEPTIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "OCTUBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "NOVIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValores (e.Cell.Row.Cells["Nombre"].Value.ToString(), "DICIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}
			#endregion Actualiza Valores
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalle);
		}
		
		#region Comentado proceso de copia periodo anterior
		//			if (this.cmbAño.Text.ToString().Trim().Length == 0)
		//			{
		//				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
		//				this.cmbAño.Focus();
		//				return;
		//			}
		//
		//			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Com_Empresa Where Periodo = {0}", Convert.ToInt32(this.cmbAño.Text))) > 0)
		//			{
		//				MessageBox.Show(string.Format("Ya esta Generado el Presupuesto para el Año {0}", Convert.ToInt32(this.cmbAño.Text)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
		//				this.btnVer_Click(sender, e);
		//				this.cmbAño.Focus();
		//				return;
		//			}
		//			
		//			using (frmCom_Copiar Copiar = new frmCom_Copiar())
		//			{				
		//				if (DialogResult.OK == Copiar.ShowDialog())
		//				{
		//					//@Periodo Int, @PE Int, @PersonalLocal Bit, @PL Int, @AsigVend Bit, @AV Int
		//					string sSQL = string.Format("Exec Com_GeneraPresupuestoEmpresaDesdePeriodoAnterior {0}, {1}, 1, {2}, 1, {3}", 
		//					 Convert.ToInt32(this.cmbAño.Text), (int)Copiar.optPE.Value, (int)Copiar.optPPL.Value, (int)Copiar.optAV.Value);
		//					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		//
		//					MessageBox.Show(string.Format("Generacion correcta"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		//
		//					this.btnVer_Click(sender, e);
		//				}
		//			}
		#endregion Comentado proceso de copia periodo anterior

		private void mnuAccionesGenerarNuevo_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio(this.txtAnio, "Ingrese el periodo", 2016, DateTime.Today.Year)) return;
			
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Com_Empresa Where Periodo = {0}", Convert.ToInt32(this.txtAnio.Value))) > 0)
			{
				MessageBox.Show(string.Format("Ya esta Generado el Presupuesto para el Año {0}", Convert.ToInt32(this.txtAnio.Value)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.btnVer_Click(sender, e);
				this.txtAnio.Focus();
				return;
			}

			if (DialogResult.Yes == MessageBox.Show(string.Format("Se va a generar:\n\nEl Presupuesto Anual Por Locales Activos en Cero\nEl Presupuesto Mensual Por Locales Activos en Cero\nEl Presupesto Anual de Personal Por Locales Activos en Cero\n\n\nEsta Seguro de Generar el Presupuesto del Año {0}", Convert.ToInt32(this.txtAnio.Value)), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				string sSQL = string.Format("Exec Com_GeneraPresupuestoEmpresa {0}", Convert.ToInt32(this.txtAnio.Value));
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				MessageBox.Show("Generación Terminada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnVer_Click(sender, e);
			}
		}

		private void mnuAccionesCopiarUltimoMes_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio(this.txtAnio, "Ingrese el periodo", 2016, DateTime.Today.Year)) return;

			string sPeriodo = Funcion.sEscalarSQL(cdsSeteoF, "Select Top 1 Periodo From Com_AsignacionDeVendedores Order By Numero Desc");

			if (DialogResult.Yes == MessageBox.Show(string.Format("Se va a copiar la siguiente información del mes de {0}\n\n\nPresupuesto por local\nPersonal por local\nAsignacion de vendedores\n\n\n¿Esta seguro de copiar la información del mes de {0}?", sPeriodo), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
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

						oCmdActualiza.CommandText = string.Format("Exec Com_CopiaMesAnterior '{0}'", sPeriodo);
						oCmdActualiza.ExecuteNonQuery();

						MessageBox.Show(string.Format("Se copio la Información del Periodo {0} correctamente", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							
						oTransaction.Commit();	

						oConexion.Close();

						oTransaction.Dispose();

						//MessageBox.Show("Nota de Crédito generada correctamente");											

						this.btnVer_Click(sender, e);
					}
					catch(Exception Exc)
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", Exc.GetType(), Exc.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);						
					}													
				}
			}
		}

		private void mnuAccionesAgregarNuevoLocal_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio(this.txtAnio, "Ingrese el periodo", 2016, DateTime.Today.Year)) return;

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Com_Empresa Where Periodo = {0}", Convert.ToInt32(this.txtAnio.Value))) == 0)
			{
				MessageBox.Show(string.Format("Para agregar un local debe generar el Presupuesto del año {0}", Convert.ToInt32(this.txtAnio.Value)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.btnVer_Click(sender, e);
				this.txtAnio.Focus();
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, 
				string.Format("Select COUNT(*) From Bodega Where Activo = 1 And Factura = 1 And Credito = 1 And Bodega Not In (Select Bodega From Com_Empresa Where Periodo = {0})", Convert.ToInt32(this.txtAnio.Value))) == 0)
			{
				MessageBox.Show("No se han agregado locales nuevos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAnio.Focus();
				return;
			}

			using (frmCom_AgregaLocales CAL = new frmCom_AgregaLocales(Convert.ToInt32(this.txtAnio.Value)))
			{				
				if (DialogResult.OK == CAL.ShowDialog())
				{
					string sSQL = string.Format("Exec Com_AgregaLocalNuevo '{0}', {1}", CAL.cmbAño.Text.ToString(), (int)CAL.cmbBodega.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);

					MessageBox.Show(string.Format("Se agrego el local '{0}' correctamente", CAL.cmbBodega.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.btnVer_Click(sender, e);
				}
			}
		}

		private void btnComandos_Click(object sender, System.EventArgs e)
		{
			this.mnuAcciones.Show(this.btnComandos, new Point(1, 25));
		}

		private void uGridDetalle_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			#region Valida Periodos Cerrados
			string sPeriodo = "";

			if (e.Cell.Column.ToString() == "ENERO")
			{
				sPeriodo = "01/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				sPeriodo = "02/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				sPeriodo = "03/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				sPeriodo = "04/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				sPeriodo = "05/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				sPeriodo = "06/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				sPeriodo = "07/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				sPeriodo = "08/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				sPeriodo = "09/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				sPeriodo = "10/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				sPeriodo = "11/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				sPeriodo = "12/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}
			#endregion Valida Periodos Cerrados
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDetalle);
		}

		private void txtAnio_Click(object sender, System.EventArgs e)
		{
			this.txtAnio.SelectAll();
		}

		private void txtAnio_Enter(object sender, System.EventArgs e)
		{
			this.txtAnio.SelectAll();
		}

		private void frmCom_PresupuestoEmpresa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridDetalleGEx_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ENERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "FEBRERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MARZO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ABRIL", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MAYO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JUNIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JULIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "AGOSTO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "SEPTIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "OCTUBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "NOVIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValoresGEx (e.Cell.Row.Cells["Nombre"].Value.ToString(), "DICIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}
			#endregion Actualiza Valores
		}

		private void uGridDetalleGEx_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			#region Valida Periodos Cerrados
			string sPeriodo = "";

			if (e.Cell.Column.ToString() == "ENERO")
			{
				sPeriodo = "01/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				sPeriodo = "02/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				sPeriodo = "03/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				sPeriodo = "04/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				sPeriodo = "05/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				sPeriodo = "06/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				sPeriodo = "07/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				sPeriodo = "08/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				sPeriodo = "09/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				sPeriodo = "10/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				sPeriodo = "11/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				sPeriodo = "12/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}
			#endregion Valida Periodos Cerrados
		}

		private void uGridDetalleGEx_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "ENERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FEBRERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MARZO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ABRIL", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MAYO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JUNIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JULIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AGOSTO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEPTIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "OCTUBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "NOVIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DICIEMBRE", 2);
		}

		private void uGridDetalleGEx_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalleGEx);
		}

		private void btnCargosDeVentas_Click(object sender, System.EventArgs e)
		{
			frmReporteCargoDeVentas CDV = new frmReporteCargoDeVentas();
			CDV.ShowDialog();
		}

		private void uGridDetalleContado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalleContado);
		}

		private void uGridDetalleContado_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			#region Valida Periodos Cerrados
			string sPeriodo = "";

			if (e.Cell.Column.ToString() == "ENERO")
			{
				sPeriodo = "01/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				sPeriodo = "02/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				sPeriodo = "03/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				sPeriodo = "04/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				sPeriodo = "05/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				sPeriodo = "06/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				sPeriodo = "07/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				sPeriodo = "08/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				sPeriodo = "09/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}				
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				sPeriodo = "10/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				sPeriodo = "11/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}				
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				sPeriodo = "12/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}				
			}
			#endregion Valida Periodos Cerrados
		}

		private void uGridDetalleContado_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "ENERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FEBRERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MARZO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ABRIL", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MAYO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JUNIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JULIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AGOSTO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEPTIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "OCTUBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "NOVIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DICIEMBRE", 2);			
		}

		private void uGridDetalleContado_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ENERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "FEBRERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MARZO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ABRIL", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MAYO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JUNIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JULIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "AGOSTO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "SEPTIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "OCTUBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "NOVIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValoresContado (e.Cell.Row.Cells["Nombre"].Value.ToString(), "DICIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}
			#endregion Actualiza Valores
		}

		private void uGridDetalleCredito_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "ENERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FEBRERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MARZO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ABRIL", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MAYO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JUNIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JULIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AGOSTO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEPTIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "OCTUBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "NOVIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DICIEMBRE", 2);
		}

		private void uGridDetalleCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalleCredito);
		}

		private void uGridDetalleCredito_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			#region Valida Periodos Cerrados
			string sPeriodo = "";

			if (e.Cell.Column.ToString() == "ENERO")
			{
				sPeriodo = "01/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				sPeriodo = "02/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				sPeriodo = "03/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				sPeriodo = "04/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				sPeriodo = "05/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				sPeriodo = "06/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				sPeriodo = "07/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				sPeriodo = "08/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				sPeriodo = "09/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}				
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				sPeriodo = "10/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				sPeriodo = "11/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}				
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				sPeriodo = "12/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}				
			}
			#endregion Valida Periodos Cerrados
		}

		private void uGridDetalleCredito_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ENERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "FEBRERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MARZO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ABRIL", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MAYO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JUNIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JULIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "AGOSTO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "SEPTIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "OCTUBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "NOVIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValoresCredito (e.Cell.Row.Cells["Nombre"].Value.ToString(), "DICIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}
			#endregion Actualiza Valores
		}

		private void uGridCarpa_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "ENERO", 0);
			FuncionesProcedimientos.FormatoGrid(e, "FEBRERO", 0);
			FuncionesProcedimientos.FormatoGrid(e, "MARZO", 0);
			FuncionesProcedimientos.FormatoGrid(e, "ABRIL", 0);
			FuncionesProcedimientos.FormatoGrid(e, "MAYO", 0);
			FuncionesProcedimientos.FormatoGrid(e, "JUNIO", 0);
			FuncionesProcedimientos.FormatoGrid(e, "JULIO", 0);
			FuncionesProcedimientos.FormatoGrid(e, "AGOSTO", 0);
			FuncionesProcedimientos.FormatoGrid(e, "SEPTIEMBRE", 0);
			FuncionesProcedimientos.FormatoGrid(e, "OCTUBRE", 0);
			FuncionesProcedimientos.FormatoGrid(e, "NOVIEMBRE", 0);
			FuncionesProcedimientos.FormatoGrid(e, "DICIEMBRE", 0);
		}

		private void uGridCarpa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridCarpa);
		}

		private void uGridCarpa_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			#region Valida Periodos Cerrados
			string sPeriodo = "";

			if (e.Cell.Column.ToString() == "ENERO")
			{
				sPeriodo = "01/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				sPeriodo = "02/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				sPeriodo = "03/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				sPeriodo = "04/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				sPeriodo = "05/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				sPeriodo = "06/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				sPeriodo = "07/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				sPeriodo = "08/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				sPeriodo = "09/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}				
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				sPeriodo = "10/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				sPeriodo = "11/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}				
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				sPeriodo = "12/" + this.txtAnio.Value.ToString();

				if (FuncionesProcedimientos.ValidaPeriodoCerrado(sPeriodo, cdsSeteoF))
				{
					MessageBox.Show("Periodo Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}				
			}
			#endregion Valida Periodos Cerrados
		}

		private void uGridCarpa_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ENERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "FEBRERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MARZO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ABRIL", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MAYO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JUNIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JULIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "AGOSTO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "SEPTIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "OCTUBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "NOVIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValoresCarpa (e.Cell.Row.Cells["Nombre"].Value.ToString(), "DICIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}
			#endregion Actualiza Valores
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.uGridDetalle);
		}

		private void menuItem2_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDetalleCredito);
		}

		private void menuItem3_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDetalleGEx);
		}

		private void menuItem4_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridCarpa);
		}

		private void BtnExcel_Click(object sender, System.EventArgs e)
		{
			this.contextMenuExportar.Show(this.BtnExcel, new Point(1, 25));
		}

		private void mnuAcciones_Popup(object sender, System.EventArgs e)
		{
		
		}

		private void contextMenuExportar_Popup(object sender, System.EventArgs e)
		{
		
		}
	}
}
