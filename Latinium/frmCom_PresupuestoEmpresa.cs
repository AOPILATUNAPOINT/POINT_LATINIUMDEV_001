using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.Text;

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
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridAccesorios;
		private System.Windows.Forms.MenuItem mnuExpAccesorios;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalleCedulas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.MenuItem menuItem5;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridMotos;
		private System.Windows.Forms.MenuItem mnuExpMotos;
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
			Infragistics.Win.Appearance appearance172 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance173 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance174 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance175 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance176 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance177 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance178 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance179 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance180 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance181 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance182 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance183 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance184 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance185 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings66 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 1, true);
			Infragistics.Win.Appearance appearance186 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings67 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 2, true);
			Infragistics.Win.Appearance appearance187 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings68 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 3, true);
			Infragistics.Win.Appearance appearance188 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings69 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 4, true);
			Infragistics.Win.Appearance appearance189 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings70 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 5, true);
			Infragistics.Win.Appearance appearance190 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings71 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 6, true);
			Infragistics.Win.Appearance appearance191 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings72 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 7, true);
			Infragistics.Win.Appearance appearance192 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings73 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 8, true);
			Infragistics.Win.Appearance appearance193 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings74 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 9, true);
			Infragistics.Win.Appearance appearance194 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings75 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 10, true);
			Infragistics.Win.Appearance appearance195 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings76 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 11, true);
			Infragistics.Win.Appearance appearance196 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings77 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance197 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings78 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 13, true);
			Infragistics.Win.Appearance appearance198 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance199 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance200 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance201 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance202 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance203 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance204 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance205 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance206 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance207 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance208 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance209 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance210 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance211 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance212 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance213 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance214 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance215 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance216 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance217 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance218 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance219 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings79 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 1, true);
			Infragistics.Win.Appearance appearance220 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings80 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 2, true);
			Infragistics.Win.Appearance appearance221 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings81 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 3, true);
			Infragistics.Win.Appearance appearance222 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings82 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 4, true);
			Infragistics.Win.Appearance appearance223 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings83 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 5, true);
			Infragistics.Win.Appearance appearance224 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings84 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 6, true);
			Infragistics.Win.Appearance appearance225 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings85 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 7, true);
			Infragistics.Win.Appearance appearance226 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings86 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 8, true);
			Infragistics.Win.Appearance appearance227 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings87 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 9, true);
			Infragistics.Win.Appearance appearance228 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings88 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 10, true);
			Infragistics.Win.Appearance appearance229 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings89 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 11, true);
			Infragistics.Win.Appearance appearance230 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings90 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance231 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings91 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 13, true);
			Infragistics.Win.Appearance appearance232 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance233 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance234 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance235 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance236 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance237 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance238 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance239 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance240 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance241 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance242 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance243 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance244 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance245 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance246 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance247 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance248 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance249 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance250 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance251 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance252 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance253 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings92 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 1, true);
			Infragistics.Win.Appearance appearance254 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings93 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 2, true);
			Infragistics.Win.Appearance appearance255 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings94 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 3, true);
			Infragistics.Win.Appearance appearance256 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings95 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 4, true);
			Infragistics.Win.Appearance appearance257 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings96 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 5, true);
			Infragistics.Win.Appearance appearance258 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings97 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 6, true);
			Infragistics.Win.Appearance appearance259 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings98 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 7, true);
			Infragistics.Win.Appearance appearance260 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings99 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 8, true);
			Infragistics.Win.Appearance appearance261 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings100 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 9, true);
			Infragistics.Win.Appearance appearance262 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings101 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 10, true);
			Infragistics.Win.Appearance appearance263 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings102 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 11, true);
			Infragistics.Win.Appearance appearance264 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings103 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 12, true);
			Infragistics.Win.Appearance appearance265 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings104 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 13, true);
			Infragistics.Win.Appearance appearance266 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance267 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance268 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance269 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance270 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance271 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance272 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance273 = new Infragistics.Win.Appearance();
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
			this.mnuExpAccesorios = new System.Windows.Forms.MenuItem();
			this.menuItem5 = new System.Windows.Forms.MenuItem();
			this.mnuExpMotos = new System.Windows.Forms.MenuItem();
			this.uGridAccesorios = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridDetalleCedulas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridMotos = new Infragistics.Win.UltraWinGrid.UltraGrid();
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
			((System.ComponentModel.ISupportInitialize)(this.uGridAccesorios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleCedulas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMotos)).BeginInit();
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
			ultraGridColumn2.Width = 74;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance3;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 74;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 74;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance5;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 74;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance6;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 74;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance7;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 74;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance8;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 75;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance9;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 76;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance10;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 78;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance11;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 76;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance12;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 76;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance13;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 76;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance14;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 81;
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
			this.uGridDetalle.Size = new System.Drawing.Size(1163, 144);
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
			this.groupBox1.Size = new System.Drawing.Size(1272, 8);
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
			this.btnComandos.Location = new System.Drawing.Point(856, 8);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.TabIndex = 726;
			this.btnComandos.Text = "Acciones";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Visible = false;
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
			this.uGridDetalleGEx.Location = new System.Drawing.Point(7, 528);
			this.uGridDetalleGEx.Name = "uGridDetalleGEx";
			this.uGridDetalleGEx.Size = new System.Drawing.Size(1163, 160);
			this.uGridDetalleGEx.TabIndex = 735;
			this.uGridDetalleGEx.Visible = false;
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
			this.btnCargosDeVentas.Location = new System.Drawing.Point(944, 8);
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
			this.uGridDetalleCredito.Location = new System.Drawing.Point(8, 192);
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
			this.uGridCarpa.Location = new System.Drawing.Point(8, 656);
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
			this.BtnExcel.Location = new System.Drawing.Point(1088, 8);
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
																																												this.menuItem4,
																																												this.mnuExpAccesorios,
																																												this.menuItem5,
																																												this.mnuExpMotos});
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
			this.menuItem3.Visible = false;
			this.menuItem3.Click += new System.EventHandler(this.menuItem3_Click);
			// 
			// menuItem4
			// 
			this.menuItem4.Index = 3;
			this.menuItem4.Text = "Carpa";
			this.menuItem4.Click += new System.EventHandler(this.menuItem4_Click);
			// 
			// mnuExpAccesorios
			// 
			this.mnuExpAccesorios.Index = 4;
			this.mnuExpAccesorios.Text = "Accesorios";
			this.mnuExpAccesorios.Click += new System.EventHandler(this.mnuExpAccesorios_Click);
			// 
			// menuItem5
			// 
			this.menuItem5.Index = 5;
			this.menuItem5.Text = "Meta Cedulas";
			this.menuItem5.Click += new System.EventHandler(this.menuItem5_Click);
			// 
			// mnuExpMotos
			// 
			this.mnuExpMotos.Index = 6;
			this.mnuExpMotos.Text = "Motos";
			this.mnuExpMotos.Click += new System.EventHandler(this.mnuExpMotos_Click);
			// 
			// uGridAccesorios
			// 
			this.uGridAccesorios.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridAccesorios.DataSource = this.ultraDataSource2;
			appearance172.BackColor = System.Drawing.Color.White;
			this.uGridAccesorios.DisplayLayout.Appearance = appearance172;
			this.uGridAccesorios.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn71.Header.Caption = "LOCAL";
			ultraGridColumn71.Header.VisiblePosition = 0;
			ultraGridColumn71.Width = 166;
			appearance173.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance173;
			ultraGridColumn72.Format = "#,##0.00";
			ultraGridColumn72.Header.VisiblePosition = 1;
			ultraGridColumn72.Width = 74;
			appearance174.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance174;
			ultraGridColumn73.Format = "#,##0.00";
			ultraGridColumn73.Header.VisiblePosition = 2;
			ultraGridColumn73.Width = 74;
			appearance175.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance175;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.VisiblePosition = 3;
			ultraGridColumn74.Width = 74;
			appearance176.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance176;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.VisiblePosition = 4;
			ultraGridColumn75.Width = 74;
			appearance177.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance177;
			ultraGridColumn76.Format = "#,##0.00";
			ultraGridColumn76.Header.VisiblePosition = 5;
			ultraGridColumn76.Width = 74;
			appearance178.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance178;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.VisiblePosition = 6;
			ultraGridColumn77.Width = 74;
			appearance179.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance179;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.VisiblePosition = 7;
			ultraGridColumn78.Width = 74;
			appearance180.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance180;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.VisiblePosition = 8;
			ultraGridColumn79.Width = 76;
			appearance181.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance181;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.VisiblePosition = 9;
			ultraGridColumn80.Width = 78;
			appearance182.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn81.CellAppearance = appearance182;
			ultraGridColumn81.Format = "#,##0.00";
			ultraGridColumn81.Header.VisiblePosition = 10;
			ultraGridColumn81.Width = 76;
			appearance183.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn82.CellAppearance = appearance183;
			ultraGridColumn82.Format = "#,##0.00";
			ultraGridColumn82.Header.VisiblePosition = 11;
			ultraGridColumn82.Width = 76;
			appearance184.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn83.CellAppearance = appearance184;
			ultraGridColumn83.Format = "#,##0.00";
			ultraGridColumn83.Header.VisiblePosition = 12;
			ultraGridColumn83.Width = 76;
			ultraGridColumn84.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance185.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn84.CellAppearance = appearance185;
			ultraGridColumn84.Format = "#,##0.00";
			ultraGridColumn84.Header.VisiblePosition = 13;
			ultraGridColumn84.Width = 76;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn84});
			ultraGridBand6.Header.Caption = "Accesorios";
			ultraGridBand6.HeaderVisible = true;
			ultraGridBand6.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance186.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings66.Appearance = appearance186;
			summarySettings66.DisplayFormat = "{0: #,##0.00}";
			summarySettings66.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance187.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings67.Appearance = appearance187;
			summarySettings67.DisplayFormat = "{0: #,##0.00}";
			summarySettings67.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance188.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings68.Appearance = appearance188;
			summarySettings68.DisplayFormat = "{0: #,##0.00}";
			summarySettings68.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance189.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings69.Appearance = appearance189;
			summarySettings69.DisplayFormat = "{0: #,##0.00}";
			summarySettings69.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance190.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings70.Appearance = appearance190;
			summarySettings70.DisplayFormat = "{0: #,##0.00}";
			summarySettings70.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance191.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings71.Appearance = appearance191;
			summarySettings71.DisplayFormat = "{0: #,##0.00}";
			summarySettings71.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance192.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings72.Appearance = appearance192;
			summarySettings72.DisplayFormat = "{0: #,##0.00}";
			summarySettings72.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance193.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings73.Appearance = appearance193;
			summarySettings73.DisplayFormat = "{0: #,##0.00}";
			summarySettings73.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance194.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings74.Appearance = appearance194;
			summarySettings74.DisplayFormat = "{0: #,##0.00}";
			summarySettings74.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance195.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings75.Appearance = appearance195;
			summarySettings75.DisplayFormat = "{0: #,##0.00}";
			summarySettings75.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance196.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings76.Appearance = appearance196;
			summarySettings76.DisplayFormat = "{0: #,##0.00}";
			summarySettings76.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance197.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings77.Appearance = appearance197;
			summarySettings77.DisplayFormat = "{0: #,##0.00}";
			summarySettings77.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance198.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings78.Appearance = appearance198;
			summarySettings78.DisplayFormat = "{0: #,##0.00}";
			summarySettings78.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand6.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings78});
			ultraGridBand6.SummaryFooterCaption = "TOTALES";
			this.uGridAccesorios.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance199.ForeColor = System.Drawing.Color.Black;
			appearance199.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAccesorios.DisplayLayout.Override.ActiveRowAppearance = appearance199;
			this.uGridAccesorios.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridAccesorios.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridAccesorios.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance200.BackColor = System.Drawing.Color.Transparent;
			this.uGridAccesorios.DisplayLayout.Override.CardAreaAppearance = appearance200;
			appearance201.ForeColor = System.Drawing.Color.Black;
			appearance201.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAccesorios.DisplayLayout.Override.CellAppearance = appearance201;
			this.uGridAccesorios.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance202.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance202.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance202.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance202.FontData.BoldAsString = "True";
			appearance202.FontData.Name = "Arial";
			appearance202.FontData.SizeInPoints = 8F;
			appearance202.ForeColor = System.Drawing.Color.White;
			appearance202.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridAccesorios.DisplayLayout.Override.HeaderAppearance = appearance202;
			appearance203.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance203.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance203.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAccesorios.DisplayLayout.Override.RowAlternateAppearance = appearance203;
			appearance204.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance204.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance204.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAccesorios.DisplayLayout.Override.RowSelectorAppearance = appearance204;
			appearance205.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance205.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance205.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAccesorios.DisplayLayout.Override.SelectedRowAppearance = appearance205;
			this.uGridAccesorios.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridAccesorios.Location = new System.Drawing.Point(8, 344);
			this.uGridAccesorios.Name = "uGridAccesorios";
			this.uGridAccesorios.Size = new System.Drawing.Size(1163, 160);
			this.uGridAccesorios.TabIndex = 741;
			this.uGridAccesorios.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridAccesorios_BeforeCellUpdate);
			this.uGridAccesorios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridAccesorios_KeyDown);
			this.uGridAccesorios.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridAccesorios_AfterCellUpdate);
			this.uGridAccesorios.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridAccesorios_InitializeLayout);
			// 
			// uGridDetalleCedulas
			// 
			this.uGridDetalleCedulas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalleCedulas.DataSource = this.ultraDataSource4;
			appearance206.BackColor = System.Drawing.Color.White;
			this.uGridDetalleCedulas.DisplayLayout.Appearance = appearance206;
			this.uGridDetalleCedulas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn85.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn85.Header.Caption = "LOCAL";
			ultraGridColumn85.Header.VisiblePosition = 0;
			ultraGridColumn85.Width = 166;
			appearance207.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn86.CellAppearance = appearance207;
			ultraGridColumn86.Format = "#,##0";
			ultraGridColumn86.Header.VisiblePosition = 1;
			ultraGridColumn86.Width = 74;
			appearance208.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn87.CellAppearance = appearance208;
			ultraGridColumn87.Format = "#,##0";
			ultraGridColumn87.Header.VisiblePosition = 2;
			ultraGridColumn87.Width = 74;
			appearance209.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn88.CellAppearance = appearance209;
			ultraGridColumn88.Format = "#,##0";
			ultraGridColumn88.Header.VisiblePosition = 3;
			ultraGridColumn88.Width = 74;
			appearance210.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn89.CellAppearance = appearance210;
			ultraGridColumn89.Format = "#,##0";
			ultraGridColumn89.Header.VisiblePosition = 4;
			ultraGridColumn89.Width = 74;
			appearance211.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn90.CellAppearance = appearance211;
			ultraGridColumn90.Format = "#,##0";
			ultraGridColumn90.Header.VisiblePosition = 5;
			ultraGridColumn90.Width = 74;
			appearance212.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn91.CellAppearance = appearance212;
			ultraGridColumn91.Format = "#,##0";
			ultraGridColumn91.Header.VisiblePosition = 6;
			ultraGridColumn91.Width = 74;
			appearance213.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn92.CellAppearance = appearance213;
			ultraGridColumn92.Format = "#,##0";
			ultraGridColumn92.Header.VisiblePosition = 7;
			ultraGridColumn92.Width = 74;
			appearance214.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn93.CellAppearance = appearance214;
			ultraGridColumn93.Format = "#,##0";
			ultraGridColumn93.Header.VisiblePosition = 8;
			ultraGridColumn93.Width = 76;
			appearance215.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn94.CellAppearance = appearance215;
			ultraGridColumn94.Format = "#,##0";
			ultraGridColumn94.Header.VisiblePosition = 9;
			ultraGridColumn94.Width = 78;
			appearance216.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn95.CellAppearance = appearance216;
			ultraGridColumn95.Format = "#,##0";
			ultraGridColumn95.Header.VisiblePosition = 10;
			ultraGridColumn95.Width = 76;
			appearance217.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn96.CellAppearance = appearance217;
			ultraGridColumn96.Format = "#,##0";
			ultraGridColumn96.Header.VisiblePosition = 11;
			ultraGridColumn96.Width = 76;
			appearance218.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn97.CellAppearance = appearance218;
			ultraGridColumn97.Format = "#,##0";
			ultraGridColumn97.Header.VisiblePosition = 12;
			ultraGridColumn97.Width = 76;
			ultraGridColumn98.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance219.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn98.CellAppearance = appearance219;
			ultraGridColumn98.Format = "#,##0";
			ultraGridColumn98.Header.VisiblePosition = 13;
			ultraGridColumn98.Width = 76;
			ultraGridBand7.Columns.AddRange(new object[] {
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
																										 ultraGridColumn98});
			ultraGridBand7.Header.Caption = "Meta Consulta de Cedulas";
			ultraGridBand7.HeaderVisible = true;
			ultraGridBand7.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance220.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings79.Appearance = appearance220;
			summarySettings79.DisplayFormat = "{0: #,##0}";
			summarySettings79.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance221.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings80.Appearance = appearance221;
			summarySettings80.DisplayFormat = "{0: #,##0}";
			summarySettings80.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance222.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings81.Appearance = appearance222;
			summarySettings81.DisplayFormat = "{0: #,##0}";
			summarySettings81.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance223.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings82.Appearance = appearance223;
			summarySettings82.DisplayFormat = "{0: #,##0}";
			summarySettings82.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance224.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings83.Appearance = appearance224;
			summarySettings83.DisplayFormat = "{0: #,##0}";
			summarySettings83.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance225.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings84.Appearance = appearance225;
			summarySettings84.DisplayFormat = "{0: #,##0}";
			summarySettings84.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance226.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings85.Appearance = appearance226;
			summarySettings85.DisplayFormat = "{0: #,##0}";
			summarySettings85.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance227.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings86.Appearance = appearance227;
			summarySettings86.DisplayFormat = "{0: #,##0}";
			summarySettings86.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance228.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings87.Appearance = appearance228;
			summarySettings87.DisplayFormat = "{0: #,##0}";
			summarySettings87.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance229.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings88.Appearance = appearance229;
			summarySettings88.DisplayFormat = "{0: #,##0}";
			summarySettings88.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance230.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings89.Appearance = appearance230;
			summarySettings89.DisplayFormat = "{0: #,##0}";
			summarySettings89.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance231.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings90.Appearance = appearance231;
			summarySettings90.DisplayFormat = "{0: #,##0}";
			summarySettings90.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance232.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings91.Appearance = appearance232;
			summarySettings91.DisplayFormat = "{0: #,##0}";
			summarySettings91.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand7.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings91});
			ultraGridBand7.SummaryFooterCaption = "TOTALES";
			this.uGridDetalleCedulas.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance233.ForeColor = System.Drawing.Color.Black;
			appearance233.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleCedulas.DisplayLayout.Override.ActiveRowAppearance = appearance233;
			this.uGridDetalleCedulas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalleCedulas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDetalleCedulas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance234.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalleCedulas.DisplayLayout.Override.CardAreaAppearance = appearance234;
			appearance235.ForeColor = System.Drawing.Color.Black;
			appearance235.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalleCedulas.DisplayLayout.Override.CellAppearance = appearance235;
			this.uGridDetalleCedulas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance236.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance236.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance236.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance236.FontData.BoldAsString = "True";
			appearance236.FontData.Name = "Arial";
			appearance236.FontData.SizeInPoints = 8F;
			appearance236.ForeColor = System.Drawing.Color.White;
			appearance236.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalleCedulas.DisplayLayout.Override.HeaderAppearance = appearance236;
			appearance237.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance237.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance237.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCedulas.DisplayLayout.Override.RowAlternateAppearance = appearance237;
			appearance238.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance238.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance238.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCedulas.DisplayLayout.Override.RowSelectorAppearance = appearance238;
			appearance239.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance239.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance239.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalleCedulas.DisplayLayout.Override.SelectedRowAppearance = appearance239;
			this.uGridDetalleCedulas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalleCedulas.Location = new System.Drawing.Point(8, 496);
			this.uGridDetalleCedulas.Name = "uGridDetalleCedulas";
			this.uGridDetalleCedulas.Size = new System.Drawing.Size(1163, 160);
			this.uGridDetalleCedulas.TabIndex = 742;
			this.uGridDetalleCedulas.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDetalleCedulas_BeforeCellUpdate);
			this.uGridDetalleCedulas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalleCedulas_KeyDown);
			this.uGridDetalleCedulas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalleCedulas_AfterCellUpdate);
			this.uGridDetalleCedulas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalleCedulas_InitializeLayout);
			// 
			// ultraDataSource4
			// 
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
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn51.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn56});
			// 
			// uGridMotos
			// 
			this.uGridMotos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridMotos.DataSource = this.ultraDataSource2;
			appearance240.BackColor = System.Drawing.Color.White;
			this.uGridMotos.DisplayLayout.Appearance = appearance240;
			this.uGridMotos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn99.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn99.Header.Caption = "LOCAL";
			ultraGridColumn99.Header.VisiblePosition = 0;
			ultraGridColumn99.Width = 166;
			appearance241.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn100.CellAppearance = appearance241;
			ultraGridColumn100.Format = "#,##0.00";
			ultraGridColumn100.Header.VisiblePosition = 1;
			ultraGridColumn100.Width = 74;
			appearance242.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn101.CellAppearance = appearance242;
			ultraGridColumn101.Format = "#,##0.00";
			ultraGridColumn101.Header.VisiblePosition = 2;
			ultraGridColumn101.Width = 74;
			appearance243.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn102.CellAppearance = appearance243;
			ultraGridColumn102.Format = "#,##0.00";
			ultraGridColumn102.Header.VisiblePosition = 3;
			ultraGridColumn102.Width = 74;
			appearance244.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn103.CellAppearance = appearance244;
			ultraGridColumn103.Format = "#,##0.00";
			ultraGridColumn103.Header.VisiblePosition = 4;
			ultraGridColumn103.Width = 74;
			appearance245.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn104.CellAppearance = appearance245;
			ultraGridColumn104.Format = "#,##0.00";
			ultraGridColumn104.Header.VisiblePosition = 5;
			ultraGridColumn104.Width = 74;
			appearance246.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn105.CellAppearance = appearance246;
			ultraGridColumn105.Format = "#,##0.00";
			ultraGridColumn105.Header.VisiblePosition = 6;
			ultraGridColumn105.Width = 74;
			appearance247.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn106.CellAppearance = appearance247;
			ultraGridColumn106.Format = "#,##0.00";
			ultraGridColumn106.Header.VisiblePosition = 7;
			ultraGridColumn106.Width = 74;
			appearance248.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn107.CellAppearance = appearance248;
			ultraGridColumn107.Format = "#,##0.00";
			ultraGridColumn107.Header.VisiblePosition = 8;
			ultraGridColumn107.Width = 76;
			appearance249.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn108.CellAppearance = appearance249;
			ultraGridColumn108.Format = "#,##0.00";
			ultraGridColumn108.Header.VisiblePosition = 9;
			ultraGridColumn108.Width = 78;
			appearance250.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn109.CellAppearance = appearance250;
			ultraGridColumn109.Format = "#,##0.00";
			ultraGridColumn109.Header.VisiblePosition = 10;
			ultraGridColumn109.Width = 76;
			appearance251.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn110.CellAppearance = appearance251;
			ultraGridColumn110.Format = "#,##0.00";
			ultraGridColumn110.Header.VisiblePosition = 11;
			ultraGridColumn110.Width = 76;
			appearance252.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn111.CellAppearance = appearance252;
			ultraGridColumn111.Format = "#,##0.00";
			ultraGridColumn111.Header.VisiblePosition = 12;
			ultraGridColumn111.Width = 76;
			ultraGridColumn112.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance253.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn112.CellAppearance = appearance253;
			ultraGridColumn112.Format = "#,##0.00";
			ultraGridColumn112.Header.VisiblePosition = 13;
			ultraGridColumn112.Width = 76;
			ultraGridBand8.Columns.AddRange(new object[] {
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
																										 ultraGridColumn112});
			ultraGridBand8.Header.Caption = "Motos";
			ultraGridBand8.HeaderVisible = true;
			ultraGridBand8.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
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
			ultraGridBand8.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings104});
			ultraGridBand8.SummaryFooterCaption = "TOTALES";
			this.uGridMotos.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			appearance267.ForeColor = System.Drawing.Color.Black;
			appearance267.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridMotos.DisplayLayout.Override.ActiveRowAppearance = appearance267;
			this.uGridMotos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridMotos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridMotos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance268.BackColor = System.Drawing.Color.Transparent;
			this.uGridMotos.DisplayLayout.Override.CardAreaAppearance = appearance268;
			appearance269.ForeColor = System.Drawing.Color.Black;
			appearance269.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridMotos.DisplayLayout.Override.CellAppearance = appearance269;
			this.uGridMotos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance270.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance270.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance270.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance270.FontData.BoldAsString = "True";
			appearance270.FontData.Name = "Arial";
			appearance270.FontData.SizeInPoints = 8F;
			appearance270.ForeColor = System.Drawing.Color.White;
			appearance270.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridMotos.DisplayLayout.Override.HeaderAppearance = appearance270;
			appearance271.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance271.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance271.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMotos.DisplayLayout.Override.RowAlternateAppearance = appearance271;
			appearance272.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance272.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance272.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMotos.DisplayLayout.Override.RowSelectorAppearance = appearance272;
			appearance273.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance273.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance273.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMotos.DisplayLayout.Override.SelectedRowAppearance = appearance273;
			this.uGridMotos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridMotos.Location = new System.Drawing.Point(8, 656);
			this.uGridMotos.Name = "uGridMotos";
			this.uGridMotos.Size = new System.Drawing.Size(1163, 160);
			this.uGridMotos.TabIndex = 743;
			this.uGridMotos.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridMotos_BeforeCellUpdate);
			this.uGridMotos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridMotos_KeyDown);
			this.uGridMotos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridMotos_AfterCellUpdate);
			this.uGridMotos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridMotos_InitializeLayout);
			// 
			// frmCom_PresupuestoEmpresa
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(1183, 597);
			this.Controls.Add(this.uGridMotos);
			this.Controls.Add(this.uGridDetalleCedulas);
			this.Controls.Add(this.uGridAccesorios);
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
			((System.ComponentModel.ISupportInitialize)(this.uGridAccesorios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalleCedulas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMotos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCom_PresupuestoEmpresa_Load(object sender, System.EventArgs e)
		{
			this.txtAnio.MinValue = 2016;
			this.txtAnio.MaxValue = DateTime.Today.Year + 1;
			this.txtAnio.Value = DateTime.Today.Year;

			MenuItem mnuSeparador = new MenuItem("-");
			MenuItem mnuCargarExcel = new MenuItem("Cargar desde Excel");
			mnuCargarExcel.Click += new EventHandler(this.mnuCargarExcel_Click);
			MenuItem mnuGenerarPlantillas = new MenuItem("Generar Plantillas CSV");
			mnuGenerarPlantillas.Click += new EventHandler(this.mnuGenerarPlantillas_Click);

			this.mnuAcciones.MenuItems.Add(mnuSeparador);
			this.mnuAcciones.MenuItems.Add(mnuCargarExcel);
			this.mnuAcciones.MenuItems.Add(mnuGenerarPlantillas);

			this.btnComandos.Visible = true;
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

			sSQL = string.Format("Exec Com_ConsultaPresupuestoAccesorios {0}", Convert.ToInt32(this.txtAnio.Value));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridAccesorios);

			sSQL = string.Format("Exec Com_ConsultaMetaConsultaCedulas {0}", Convert.ToInt32(this.txtAnio.Value));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDetalleCedulas);
			
			sSQL = string.Format("Exec Com_ConsultaPresupuestoMotos {0}", Convert.ToInt32(this.txtAnio.Value));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridMotos);

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

		private void ActualizaValoresAccesorios (string sLocal, string sMes, decimal dValor)
		{
			string sSQL = string.Format("Exec Com_ActualizaValoresPresupuestoAccesorios {0}, '{1}', '{2}', {3}", 
				Convert.ToInt32(this.txtAnio.Value), sLocal, sMes, dValor);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void ActualizaValoresMotos (string sLocal, string sMes, decimal dValor)
		{
			string sSQL = string.Format("Exec Com_ActualizaValoresPresupuestoMotos {0}, '{1}', '{2}', {3}", 
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
		private void ActualizaValoresConsultaCedulas (string sLocal, string sMes, decimal dValor)
		{
			string sSQL = string.Format("Exec Com_ActualizaValoresConsultaCedulas {0}, '{1}', '{2}', {3}", 
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

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridAccesorios.Rows.All)
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

		private void mnuExpAccesorios_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridAccesorios);
		}

		private void uGridAccesorios_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ENERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "FEBRERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MARZO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ABRIL", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MAYO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JUNIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JULIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "AGOSTO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "SEPTIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "OCTUBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "NOVIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValoresAccesorios (e.Cell.Row.Cells["Nombre"].Value.ToString(), "DICIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}
			#endregion Actualiza Valores
		}

		private void uGridAccesorios_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
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

		private void uGridAccesorios_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
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

		private void uGridAccesorios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridAccesorios);
		}

		private void menuItem5_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDetalleCedulas);
		}

		private void uGridDetalleCedulas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValoresConsultaCedulas (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ENERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValoresConsultaCedulas (e.Cell.Row.Cells["Nombre"].Value.ToString(), "FEBRERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValoresConsultaCedulas (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MARZO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValoresConsultaCedulas(e.Cell.Row.Cells["Nombre"].Value.ToString(), "ABRIL", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValoresConsultaCedulas (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MAYO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValoresConsultaCedulas (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JUNIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValoresConsultaCedulas (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JULIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValoresConsultaCedulas (e.Cell.Row.Cells["Nombre"].Value.ToString(), "AGOSTO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValoresConsultaCedulas (e.Cell.Row.Cells["Nombre"].Value.ToString(), "SEPTIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValoresConsultaCedulas (e.Cell.Row.Cells["Nombre"].Value.ToString(), "OCTUBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValoresConsultaCedulas (e.Cell.Row.Cells["Nombre"].Value.ToString(), "NOVIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValoresConsultaCedulas (e.Cell.Row.Cells["Nombre"].Value.ToString(), "DICIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}
			#endregion Actualiza Valores
		}

		private void uGridDetalleCedulas_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
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

		private void uGridDetalleCedulas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalleCedulas);
		}

		private void uGridDetalleCedulas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridMotos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
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

		private void uGridMotos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Actualiza Valores
			if (e.Cell.Column.ToString() == "ENERO")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ENERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "FEBRERO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MARZO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "ABRIL", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "MAYO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JUNIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "JULIO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "AGOSTO", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "SEPTIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "OCTUBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "NOVIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				ActualizaValoresMotos (e.Cell.Row.Cells["Nombre"].Value.ToString(), "DICIEMBRE", Convert.ToDecimal(e.Cell.Value));
				Sumatorias(e);
			}
			#endregion Actualiza Valores
		}

		private void uGridMotos_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
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

		private void uGridMotos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridMotos);
		}

		private void mnuExpMotos_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridMotos);
		}

		#region Importar desde Excel

		private delegate decimal ValorSelector(FilaMeta f);
		private static decimal GetGeneral(FilaMeta f)    { return f.General; }
		private static decimal GetCredito(FilaMeta f)    { return f.Credito; }
		private static decimal GetAccesorios(FilaMeta f) { return f.Accesorios; }

		// Estructura interna para cada fila del Excel
		private class FilaMeta
		{
			public string Local;
			public decimal General;
			public decimal Credito;
			public decimal Accesorios;
		}

		private void mnuCargarExcel_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio(this.txtAnio, "Ingrese el periodo", 2016, DateTime.Today.Year + 1)) return;

			OpenFileDialog dlg = new OpenFileDialog();
			dlg.Filter = "Excel (*.xlsx;*.xls)|*.xlsx;*.xls|CSV (*.csv)|*.csv";
			dlg.Title = "Seleccionar archivo de Metas";
			if (dlg.ShowDialog() != DialogResult.OK) return;

			string sArchivo = dlg.FileName;
			string sExt = Path.GetExtension(sArchivo).ToLower();

			DataTable dt = new DataTable();
			try
			{
				if (sExt == ".csv")
				{
					dt = LeerCSV(sArchivo);
				}
				else
				{
					string provider = sExt == ".xlsx"
						? "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=NO;IMEX=1';"
						: "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
					string strConn = string.Format(provider, sArchivo);
					OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Hoja1$]", strConn);
					da.Fill(dt);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al leer el archivo:\n" + ex.Message, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			// Detectar fila de encabezado (donde col[0] == "LOCAL") y el mes
			int iHeaderRow = -1;
			string sMes = "";
			string[] meses = { "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE" };

			for (int i = 0; i < dt.Rows.Count; i++)
			{
				string col0 = dt.Rows[i][0] == DBNull.Value ? "" : dt.Rows[i][0].ToString().Trim().ToUpper();
				if (col0 == "LOCAL")
				{
					iHeaderRow = i;
					// Buscar el mes en los encabezados de columnas siguientes
					for (int c = 1; c < dt.Columns.Count && c <= 4; c++)
					{
						string hdr = dt.Rows[i][c] == DBNull.Value ? "" : dt.Rows[i][c].ToString().ToUpper();
						foreach (string m in meses)
						{
							if (hdr.IndexOf(m) >= 0) { sMes = m; break; }
						}
						if (sMes.Length > 0) break;
					}
					break;
				}
			}

			if (iHeaderRow < 0)
			{
				MessageBox.Show("No se encontró la fila de encabezados (LOCAL).\nVerifique que el archivo tenga el formato correcto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			// Si no se detectó el mes, preguntar al usuario
			if (sMes.Length == 0)
			{
				sMes = SolicitarMes();
				if (sMes.Length == 0) return;
			}

			// Parsear filas de datos
			ArrayList datos = new ArrayList();
			for (int i = iHeaderRow + 1; i < dt.Rows.Count; i++)
			{
				string local = dt.Rows[i][0] == DBNull.Value ? "" : dt.Rows[i][0].ToString().Trim();
				if (local.Length == 0) continue;
				if (local.ToUpper().StartsWith("TOTAL") || local.ToUpper().StartsWith(" TOTAL")) continue;

				FilaMeta f = new FilaMeta();
				f.Local    = local;
				f.General   = ParsearDecimalExcel(dt.Rows[i].ItemArray.Length > 1 ? dt.Rows[i][1] : DBNull.Value);
				f.Credito   = ParsearDecimalExcel(dt.Rows[i].ItemArray.Length > 2 ? dt.Rows[i][2] : DBNull.Value);
				f.Accesorios = ParsearDecimalExcel(dt.Rows[i].ItemArray.Length > 3 ? dt.Rows[i][3] : DBNull.Value);
				datos.Add(f);
			}

			if (datos.Count == 0)
			{
				MessageBox.Show("No se encontraron datos en el archivo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			// Confirmación por tabla
			if (!ConfirmarCarga("META GENERAL",    sMes, datos, new ValorSelector(GetGeneral)))    return;
			if (!ConfirmarCarga("META CREDITO",    sMes, datos, new ValorSelector(GetCredito)))    return;
			if (!ConfirmarCarga("META ACCESORIOS", sMes, datos, new ValorSelector(GetAccesorios))) return;

			// Ejecutar carga
			int iAnio = Convert.ToInt32(this.txtAnio.Value);
			int okGen = 0, okCred = 0, okAcc = 0;
			ArrayList noEncontrados = new ArrayList();

			foreach (FilaMeta f in datos)
			{
				// Verificar que el local existe en la BD
				int iExiste = Funcion.iEscalarSQL(cdsSeteoF,
					string.Format("SELECT COUNT(*) FROM Bodega WHERE LTRIM(RTRIM(Nombre)) = '{0}'", f.Local.Replace("'", "''")));

				if (iExiste == 0)
				{
					bool yaMarcado = false;
					foreach (string sn in noEncontrados) if (sn == f.Local) { yaMarcado = true; break; }
					if (!yaMarcado)
						noEncontrados.Add(f.Local);
					continue;
				}

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("EXEC Com_ActualizaValoresPresupuestoEmpresa {0}, '{1}', '{2}', {3}",
					iAnio, f.Local.Replace("'", "''"), sMes, f.General.ToString(System.Globalization.CultureInfo.InvariantCulture)));
				okGen++;

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("EXEC Com_ActualizaValoresPresupuestoEmpresaCredito {0}, '{1}', '{2}', {3}",
					iAnio, f.Local.Replace("'", "''"), sMes, f.Credito.ToString(System.Globalization.CultureInfo.InvariantCulture)));
				okCred++;

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("EXEC Com_ActualizaValoresPresupuestoAccesorios {0}, '{1}', '{2}', {3}",
					iAnio, f.Local.Replace("'", "''"), sMes, f.Accesorios.ToString(System.Globalization.CultureInfo.InvariantCulture)));
				okAcc++;
			}

			// Resumen
			StringBuilder sbResumen = new StringBuilder();
			sbResumen.Append("========== RESUMEN DE CARGA ==========").Append("\r\n");
			sbResumen.Append(string.Format("Periodo : {0} - {1}", sMes, iAnio)).Append("\r\n");
			sbResumen.Append(string.Format("Archivo : {0}", Path.GetFileName(sArchivo))).Append("\r\n");
			sbResumen.Append("").Append("\r\n");
			sbResumen.Append(string.Format("Meta General     : {0} locales actualizados", okGen)).Append("\r\n");
			sbResumen.Append(string.Format("Meta Crédito     : {0} locales actualizados", okCred)).Append("\r\n");
			sbResumen.Append(string.Format("Meta Accesorios  : {0} locales actualizados", okAcc)).Append("\r\n");

			if (noEncontrados.Count > 0)
			{
				sbResumen.Append("").Append("\r\n");
				sbResumen.Append(string.Format("ADVERTENCIA: {0} local(es) no encontrado(s) en la BD:", noEncontrados.Count)).Append("\r\n");
				foreach (string s in noEncontrados)
					sbResumen.Append("  - " + s.ToString()).Append("\r\n");
				sbResumen.Append("").Append("\r\n");
				sbResumen.Append("Verifique que el nombre coincida exactamente con la tabla Bodega.").Append("\r\n");
			}

			MessageBox.Show(sbResumen.ToString(), "Carga Completada", MessageBoxButtons.OK,
				noEncontrados.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);

			// Refrescar grids
			this.btnVer_Click(sender, e);
		}

		private bool ConfirmarCarga(string sTabla, string sMes, ArrayList datos, ValorSelector getValor)
		{
			StringBuilder sb = new StringBuilder();
			sb.Append(string.Format("TABLA: {0}   -   MES: {1}   -   AÑO: {2}", sTabla, sMes, Convert.ToInt32(this.txtAnio.Value))).Append("\r\n");
			sb.Append(new string('=', 55)).Append("\r\n");
			sb.Append(string.Format("{0,-28} {1,14}", "LOCAL", "VALOR")).Append("\r\n");
			sb.Append(new string('-', 44)).Append("\r\n");
			decimal total = 0;
			foreach (FilaMeta f in datos)
			{
				decimal v = getValor((FilaMeta)f);
				total += v;
				sb.Append(string.Format("{0,-28} {1,14:N2}", ((FilaMeta)f).Local, v)).Append("\r\n");
			}
			sb.Append(new string('-', 44)).Append("\r\n");
			sb.Append(string.Format("{0,-28} {1,14:N2}", "TOTAL", total)).Append("\r\n");
			sb.Append("").Append("\r\n");
			sb.Append("¿Confirma la carga de esta tabla?").Append("\r\n");

			return DialogResult.Yes == MessageBox.Show(sb.ToString(),
				"Confirmar - " + sTabla, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
		}

		private decimal ParsearDecimalExcel(object valor)
		{
			if (valor == DBNull.Value || valor == null) return 0;
			if (valor is double || valor is float || valor is decimal)
				return Convert.ToDecimal(valor);

			string s = valor.ToString().Trim().Replace(" ", "");
			if (s.Length == 0) return 0;

			// Si tiene punto y coma: coma=miles, punto=decimal  ("30,780.00")
			if (s.IndexOf(".") >= 0 && s.IndexOf(",") >= 0)
				s = s.Replace(",", "");
			// Solo coma sin punto: podria ser miles ("54,000") o decimal ("1,5")
			else if (s.IndexOf(",") >= 0 && s.IndexOf(".") < 0)
			{
				int posUltimaComa = s.LastIndexOf(',');
				int digDespues = s.Length - posUltimaComa - 1;
				if (digDespues == 3) s = s.Replace(",", ""); // miles
				else s = s.Replace(",", ".");                 // decimal
			}

			decimal resultado = 0;
			try
			{
				resultado = decimal.Parse(s, System.Globalization.NumberStyles.Any,
					System.Globalization.CultureInfo.InvariantCulture);
			}
			catch { resultado = 0; }
			return resultado;
		}

		private DataTable LeerCSV(string sArchivo)
		{
			DataTable dt = new DataTable();
			ArrayList lstLineas = new ArrayList();
			StreamReader sr = new StreamReader(sArchivo, Encoding.UTF8);
			try
			{
				string sLinea;
				while ((sLinea = sr.ReadLine()) != null)
					lstLineas.Add(sLinea);
			}
			finally
			{
				sr.Close();
			}
			string[] lineas = (string[])lstLineas.ToArray(typeof(string));
			if (lineas.Length == 0) return dt;

			// Detectar el delimitador (tab, punto y coma o coma) según cuál produzca más columnas
			char[] posiblesDelimitadores = new char[] { '\t', ';', ',' };
			char delimitador = '\t';
			int maxColsDelim = 0;
			foreach (char d in posiblesDelimitadores)
			{
				int n = lineas[0].Split(d).Length;
				if (n > maxColsDelim)
				{
					maxColsDelim = n;
					delimitador = d;
				}
			}

			// Determinar número máximo de columnas
			int maxCols = 0;
			foreach (string l in lineas)
			{
				int n = l.Split(delimitador).Length;
				if (n > maxCols) maxCols = n;
			}
			for (int c = 0; c < maxCols; c++) dt.Columns.Add("F" + c);

			foreach (string linea in lineas)
			{
				string[] partes = linea.Split(delimitador);
				DataRow row = dt.NewRow();
				for (int c = 0; c < partes.Length && c < maxCols; c++)
					row[c] = partes[c].Trim().Trim('"');
				dt.Rows.Add(row);
			}
			return dt;
		}

		private string SolicitarMes()
		{
			Form fMes = new Form();
			fMes.Text = "Seleccionar Mes a Cargar";
			fMes.Size = new Size(270, 130);
			fMes.StartPosition = FormStartPosition.CenterParent;
			fMes.FormBorderStyle = FormBorderStyle.FixedDialog;
			fMes.MaximizeBox = false;
			fMes.MinimizeBox = false;

			Label lbl = new Label();
			lbl.Text = "Mes:";
			lbl.Location = new Point(15, 18);
			lbl.Size = new Size(35, 20);

			ComboBox cbo = new ComboBox();
			cbo.Location = new Point(55, 15);
			cbo.Size = new Size(185, 23);
			cbo.DropDownStyle = ComboBoxStyle.DropDownList;
			string[] meses = { "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE" };
			foreach (string m in meses) cbo.Items.Add(m);
			cbo.SelectedIndex = DateTime.Today.Month - 1;

			Button btnOk = new Button();
			btnOk.Text = "Aceptar";
			btnOk.Location = new Point(95, 55);
			btnOk.Size = new Size(80, 27);
			btnOk.DialogResult = DialogResult.OK;

			fMes.Controls.AddRange(new Control[] { lbl, cbo, btnOk });
			fMes.AcceptButton = btnOk;

			return fMes.ShowDialog() == DialogResult.OK ? cbo.SelectedItem.ToString() : "";
		}

		private void mnuGenerarPlantillas_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio(this.txtAnio, "Ingrese el periodo", 2016, DateTime.Today.Year + 1)) return;

			FolderBrowserDialog dlgCarpeta = new FolderBrowserDialog();
			dlgCarpeta.Description = "Seleccione la carpeta donde guardar las plantillas";
			dlgCarpeta.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
			if (dlgCarpeta.ShowDialog() != DialogResult.OK) return;

			string sCarpeta = dlgCarpeta.SelectedPath;
			int iAnio = Convert.ToInt32(this.txtAnio.Value);

			// Obtener locales activos
			string sSQL = string.Format(
				"SELECT LTRIM(RTRIM(Nombre)) FROM Bodega WHERE Activo=1 AND Factura=1 AND Credito=1 AND CONVERT(Date, ISNULL(FechaCierre,'20301231')) > '{0}0101' ORDER BY Nombre",
				iAnio);
			ArrayList locales = new ArrayList();
			try
			{
				cdsSeteoF.Schema.Connections[0].Open();
				string sConnStr = cdsSeteoF.Schema.Connections[0].ConnectionString;
				cdsSeteoF.Schema.Connections[0].Close();
				DataSet dsBod = new DataSet();
				SqlDataAdapter daBod = new SqlDataAdapter(sSQL, sConnStr);
				daBod.Fill(dsBod);
				foreach (DataRow r in dsBod.Tables[0].Rows)
					locales.Add(r[0].ToString());
			}
			catch { }

			if (locales.Count == 0)
			{
				MessageBox.Show("No se encontraron locales activos para el periodo " + iAnio, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			string[] sMeses = { "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE" };
			string encabezado = "LOCAL\t" + string.Join("\t", sMeses);

			// Plantillas individuales por tabla
			string[][] tablas = new string[][]
			{
				new string[] { "Plantilla_General",    "META GENERAL" },
				new string[] { "Plantilla_Credito",    "META CREDITO" },
				new string[] { "Plantilla_Accesorios", "META ACCESORIOS" },
				new string[] { "Plantilla_GEx",        "META GARANTIA EXTENDIDA" },
				new string[] { "Plantilla_Contado",    "META CONTADO" },
				new string[] { "Plantilla_Motos",      "META MOTOS" },
				new string[] { "Plantilla_Carpa",      "META CARPA" },
				new string[] { "Plantilla_Cedulas",    "META CONSULTA CEDULAS" }
			};

			ArrayList archivosCreados = new ArrayList();
			foreach (string[] t in tablas)
			{
				string sNombreArch = string.Format("{0}_{1}.csv", t[0], iAnio);
				string sRuta = Path.Combine(sCarpeta, sNombreArch);
				StringBuilder sb = new StringBuilder();
				sb.Append(t[1] + " " + iAnio).Append("\r\n");
				sb.Append(encabezado).Append("\r\n");
				foreach (string local in locales)
					sb.Append(local + "\t" + new string('\t', sMeses.Length - 1)).Append("\r\n");
				StreamWriter swPlant = new StreamWriter(sRuta, false, Encoding.UTF8);
				try { swPlant.Write(sb.ToString()); }
				finally { swPlant.Close(); }
				archivosCreados.Add(sNombreArch);
			}

			// Plantilla combinada (formato del Excel original)
			string sNombreComb = string.Format("Plantilla_Combinada_General_Credito_Accesorios_{0}.csv", iAnio);
			string sRutaComb = Path.Combine(sCarpeta, sNombreComb);
			StringBuilder sbComb = new StringBuilder();
			sbComb.Append("LOCAL\tMETA GENERAL\tMETA CREDITO\tMETA ACCESORIOS").Append("\r\n");
			foreach (string local in locales)
				sbComb.Append(local.ToString() + "\t0\t0\t0").Append("\r\n");
			StreamWriter swComb = new StreamWriter(sRutaComb, false, Encoding.UTF8);
			try { swComb.Write(sbComb.ToString()); }
			finally { swComb.Close(); }
			archivosCreados.Add(sNombreComb);

			string[] nombresArr = (string[])archivosCreados.ToArray(typeof(string));
			MessageBox.Show(
				string.Format("{0} plantillas generadas en:\n{1}\n\nArchivos:\n- {2}",
					archivosCreados.Count, sCarpeta, string.Join("\n- ", nombresArr)),
				"Plantillas Generadas", MessageBoxButtons.OK, MessageBoxIcon.Information);

			System.Diagnostics.Process.Start("explorer.exe", sCarpeta);
		}

		#endregion Importar desde Excel
	}
}
