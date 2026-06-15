using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Latinium.Models;
using Latinium.Services;
using System.Data;            // necesario para DataTable y DataRow
using Infragistics.Win.UltraWinGrid;
using Latinium.Models.FamilyFriends.Estado;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFamilyFriendsReporteria.
	/// </summary>
	public class frmFamilyFriendsReporteria : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.Button button2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsResponsable;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo1;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsFamilyFriends_ReporteCon;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArea;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		public Infragistics.Win.UltraWinGrid.UltraGrid udgSeguimiento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsInvitadoDet;
		public Infragistics.Win.UltraWinGrid.UltraGrid ugdInvitadoDetalle;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblTipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFFCabecera;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstadoFF;
		private System.ComponentModel.IContainer components;

		public frmFamilyFriendsReporteria()
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
		/// Limpiar los recursos que se están utilizando.
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFamilyFriendsReporteria));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idWEB_Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("identificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Colaborador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EmpresaEmpleado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AreaNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DepartamentoNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargoNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalInvitados");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalFacturas");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalMontoFacturas");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalNotasCreditoCount");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalNotasCredito");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalNeto", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalCosto");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenUtilidad");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre_Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeguimientoAnadir", 0);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 2);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalNeto", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalNeto", 12, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalInvitados", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalInvitados", 7, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalFacturas", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalFacturas", 8, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalMontoFacturas", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalMontoFacturas", 9, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalNotasCreditoCount", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalNotasCreditoCount", 10, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalNotasCredito", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalNotasCredito", 11, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalCosto", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalCosto", 13, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MargenUtilidad", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MargenUtilidad", 14, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idWEB_Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("identificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Colaborador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EmpresaEmpleado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AreaNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DepartamentoNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CargoNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalInvitados");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalFacturas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalMontoFacturas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalNotasCreditoCount");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalNotasCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalNeto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre_Bodega");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idWEB_FF_Invitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Invitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AreaNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PatrocinadorNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PatrocinadorCedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDocumento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoTotalUltimo");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenUtilidad");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FF");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeguimientoAnadir", 0);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 2);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Monto", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Monto", 11, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CostoTotalUltimo", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CostoTotalUltimo", 12, true);
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MargenUtilidad", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MargenUtilidad", 13, true);
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cantidad", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cantidad", 20, true);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DescuentoPorc", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DescuentoPorc", 21, true);
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubTotal", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubTotal", 22, true);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idWEB_FF_Invitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Invitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Email");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AreaNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PatrocinadorNombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PatrocinadorCedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDocumento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDocumento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDocumento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoTotalUltimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FF");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.label6 = new System.Windows.Forms.Label();
			this.udgSeguimiento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsFamilyFriends_ReporteCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblContador = new System.Windows.Forms.Label();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.button3 = new System.Windows.Forms.Button();
			this.label5 = new System.Windows.Forms.Label();
			this.ugdInvitadoDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsInvitadoDet = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.button2 = new System.Windows.Forms.Button();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.ultraCombo1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label12 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.udsResponsable = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbEstadoFF = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbFFCabecera = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblTipo = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbArea = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			this.groupBox4.SuspendLayout();
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgSeguimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsFamilyFriends_ReporteCon)).BeginInit();
			this.tabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdInvitadoDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsInvitadoDet)).BeginInit();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadoFF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFFCabecera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArea)).BeginInit();
			this.SuspendLayout();
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1048, 16);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 22);
			this.btnSalir.TabIndex = 1058;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Controls.Add(this.label1);
			this.groupBox3.Controls.Add(this.cmbResponsable);
			this.groupBox3.Location = new System.Drawing.Point(13, 14);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(982, 56);
			this.groupBox3.TabIndex = 1056;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Datos principales";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 16);
			this.label1.TabIndex = 1038;
			this.label1.Text = "* Usuario";
			// 
			// cmbResponsable
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance1;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(104, 24);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.ReadOnly = true;
			this.cmbResponsable.Size = new System.Drawing.Size(248, 21);
			this.cmbResponsable.TabIndex = 1042;
			this.cmbResponsable.ValueMember = "idPersonal";
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(888, 24);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 1019;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.CausesValidation = false;
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(1016, 8);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(78, 23);
			this.button1.TabIndex = 1057;
			this.button1.Text = "&Exportar";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// groupBox4
			// 
			this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox4.Controls.Add(this.tabControl1);
			this.groupBox4.Controls.Add(this.button2);
			this.groupBox4.Controls.Add(this.groupBox5);
			this.groupBox4.Location = new System.Drawing.Point(13, 176);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1109, 368);
			this.groupBox4.TabIndex = 1057;
			this.groupBox4.TabStop = false;
			// 
			// tabControl1
			// 
			this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Controls.Add(this.tabPage2);
			this.tabControl1.Location = new System.Drawing.Point(0, 8);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(1107, 360);
			this.tabControl1.TabIndex = 1060;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.label6);
			this.tabPage1.Controls.Add(this.udgSeguimiento);
			this.tabPage1.Controls.Add(this.lblContador);
			this.tabPage1.Controls.Add(this.button1);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(1099, 334);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "General";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.ForeColor = System.Drawing.Color.Firebrick;
			this.label6.Location = new System.Drawing.Point(13, 7);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(0, 19);
			this.label6.TabIndex = 1063;
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// udgSeguimiento
			// 
			this.udgSeguimiento.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgSeguimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgSeguimiento.DataSource = this.udsFamilyFriends_ReporteCon;
			appearance2.BackColor = System.Drawing.Color.White;
			this.udgSeguimiento.DisplayLayout.Appearance = appearance2;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Fixed = true;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 135;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Empresa";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 251;
			ultraGridColumn7.AutoEdit = false;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Area";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 223;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Departamento";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 227;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn9.Header.Caption = "Cargo";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 207;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance3;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellButtonAppearance = appearance4;
			ultraGridColumn10.Format = "";
			ultraGridColumn10.Header.Caption = "Cantidad Invitados";
			ultraGridColumn10.Header.VisiblePosition = 11;
			ultraGridColumn10.Width = 107;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance5;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellButtonAppearance = appearance6;
			ultraGridColumn11.Format = "";
			ultraGridColumn11.Header.Caption = "Total Cantidad Facturas";
			ultraGridColumn11.Header.VisiblePosition = 12;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance7;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellButtonAppearance = appearance8;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Total Monto Facturas";
			ultraGridColumn12.Header.VisiblePosition = 13;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance9;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellButtonAppearance = appearance10;
			ultraGridColumn13.Format = "";
			ultraGridColumn13.Header.Caption = "Total Cantidad Nota Crédito";
			ultraGridColumn13.Header.VisiblePosition = 14;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance11;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellButtonAppearance = appearance12;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Total Monto Notas Credito";
			ultraGridColumn14.Header.VisiblePosition = 15;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance13;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellButtonAppearance = appearance14;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Costo Total Ultimo";
			ultraGridColumn15.Header.VisiblePosition = 17;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance15;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellButtonAppearance = appearance16;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "SubTotal";
			ultraGridColumn16.Header.VisiblePosition = 18;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance17;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellButtonAppearance = appearance18;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Total Margen Utilidad";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn18.Header.Caption = "Bodega";
			ultraGridColumn18.Header.VisiblePosition = 9;
			ultraGridColumn18.Width = 140;
			appearance19.Image = ((object)(resources.GetObject("appearance19.Image")));
			ultraGridColumn19.CellAppearance = appearance19;
			appearance20.Image = ((object)(resources.GetObject("appearance20.Image")));
			ultraGridColumn19.CellButtonAppearance = appearance20;
			ultraGridColumn19.Header.Caption = "...";
			ultraGridColumn19.Header.Fixed = true;
			ultraGridColumn19.Header.VisiblePosition = 6;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 22;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.Caption = "Estado";
			ultraGridColumn20.Header.VisiblePosition = 10;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn20.Width = 122;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn21.Header.Caption = "Nº";
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Width = 29;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn21});
			ultraGridBand2.Header.Caption = "General";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance21;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance22;
			summarySettings2.DisplayFormat = "{0:#,##0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance23;
			summarySettings3.DisplayFormat = "{0:#,##0}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance24;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance25;
			summarySettings5.DisplayFormat = "{0:#,##0}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance26;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8});
			ultraGridBand2.SummaryFooterCaption = "Sum:";
			this.udgSeguimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.udgSeguimiento.DisplayLayout.GroupByBox.Hidden = true;
			this.udgSeguimiento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgSeguimiento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgSeguimiento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance27.BackColor = System.Drawing.Color.Transparent;
			this.udgSeguimiento.DisplayLayout.Override.CardAreaAppearance = appearance27;
			this.udgSeguimiento.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 10F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgSeguimiento.DisplayLayout.Override.HeaderAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgSeguimiento.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgSeguimiento.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.udgSeguimiento.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgSeguimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.udgSeguimiento.Location = new System.Drawing.Point(0, 40);
			this.udgSeguimiento.Name = "udgSeguimiento";
			this.udgSeguimiento.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.udgSeguimiento.Size = new System.Drawing.Size(1104, 296);
			this.udgSeguimiento.TabIndex = 1060;
			this.udgSeguimiento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.udgSeguimiento_InitializeLayout);
			// 
			// udsFamilyFriends_ReporteCon
			// 
			ultraDataColumn4.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn5.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn7.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			this.udsFamilyFriends_ReporteCon.Band.Columns.AddRange(new object[] {
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
																																						ultraDataColumn16});
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(8, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 1045;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.button3);
			this.tabPage2.Controls.Add(this.label5);
			this.tabPage2.Controls.Add(this.ugdInvitadoDetalle);
			this.tabPage2.Location = new System.Drawing.Point(4, 22);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Size = new System.Drawing.Size(1099, 334);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "Detalle";
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button3.CausesValidation = false;
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button3.Location = new System.Drawing.Point(1013, 7);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(78, 23);
			this.button3.TabIndex = 1063;
			this.button3.Text = "&Exportar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(8, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(0, 19);
			this.label5.TabIndex = 1062;
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ugdInvitadoDetalle
			// 
			this.ugdInvitadoDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdInvitadoDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdInvitadoDetalle.DataSource = this.udsInvitadoDet;
			appearance31.BackColor = System.Drawing.Color.White;
			this.ugdInvitadoDetalle.DisplayLayout.Appearance = appearance31;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn27.AutoEdit = false;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.Caption = "Area";
			ultraGridColumn27.Header.VisiblePosition = 7;
			ultraGridColumn27.Width = 223;
			ultraGridColumn28.Header.Caption = "Patrocinador Nombre";
			ultraGridColumn28.Header.VisiblePosition = 6;
			ultraGridColumn29.Header.Caption = "Patrocinador Cedula";
			ultraGridColumn29.Header.VisiblePosition = 8;
			ultraGridColumn30.Header.Caption = "Tipo Documento";
			ultraGridColumn30.Header.VisiblePosition = 9;
			ultraGridColumn31.Header.Caption = "Numero Documento";
			ultraGridColumn31.Header.VisiblePosition = 10;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn32.Header.Caption = "Fecha Documento";
			ultraGridColumn32.Header.VisiblePosition = 13;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance32;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellButtonAppearance = appearance33;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.Caption = "Precio Producto";
			ultraGridColumn33.Header.VisiblePosition = 20;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellButtonAppearance = appearance34;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.Caption = "Costo Total Ultimo";
			ultraGridColumn34.Header.VisiblePosition = 24;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance35;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellButtonAppearance = appearance36;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.Caption = "Margen Utilidad";
			ultraGridColumn35.Header.VisiblePosition = 17;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.Caption = "Bodega Documento";
			ultraGridColumn36.Header.VisiblePosition = 14;
			ultraGridColumn37.Header.Caption = "Forma Pago";
			ultraGridColumn37.Header.VisiblePosition = 11;
			ultraGridColumn38.Header.VisiblePosition = 15;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.Caption = "Estados FF";
			ultraGridColumn39.Header.VisiblePosition = 16;
			ultraGridColumn40.Header.VisiblePosition = 19;
			ultraGridColumn41.Header.Caption = "Codigo Producto";
			ultraGridColumn41.Header.VisiblePosition = 18;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance37;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellButtonAppearance = appearance38;
			ultraGridColumn42.Header.Caption = "Cantidad Producto";
			ultraGridColumn42.Header.VisiblePosition = 21;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance39;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellButtonAppearance = appearance40;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.Caption = "Descuento Documento";
			ultraGridColumn43.Header.VisiblePosition = 23;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance41;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellButtonAppearance = appearance42;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 25;
			appearance43.Image = ((object)(resources.GetObject("appearance43.Image")));
			ultraGridColumn45.CellAppearance = appearance43;
			appearance44.Image = ((object)(resources.GetObject("appearance44.Image")));
			ultraGridColumn45.CellButtonAppearance = appearance44;
			ultraGridColumn45.Header.Caption = "...";
			ultraGridColumn45.Header.Fixed = true;
			ultraGridColumn45.Header.VisiblePosition = 22;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn45.Width = 22;
			ultraGridColumn46.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn46.Header.Caption = "Estado";
			ultraGridColumn46.Header.VisiblePosition = 12;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn46.Width = 122;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn47.Header.Caption = "Nº";
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn47.Width = 29;
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
																										 ultraGridColumn47});
			ultraGridBand3.Header.Caption = "Detalle";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance45;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance46;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Left;
			summarySettings11.Appearance = appearance47;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.Hidden = true;
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance48;
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance49;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance50;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13,
																																															summarySettings14});
			ultraGridBand3.SummaryFooterCaption = "Sum Detalle:";
			this.ugdInvitadoDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ugdInvitadoDetalle.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdInvitadoDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugdInvitadoDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugdInvitadoDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance51.BackColor = System.Drawing.Color.Transparent;
			this.ugdInvitadoDetalle.DisplayLayout.Override.CardAreaAppearance = appearance51;
			this.ugdInvitadoDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance52.FontData.BoldAsString = "True";
			appearance52.FontData.Name = "Arial";
			appearance52.FontData.SizeInPoints = 10F;
			appearance52.ForeColor = System.Drawing.Color.White;
			appearance52.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdInvitadoDetalle.DisplayLayout.Override.HeaderAppearance = appearance52;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdInvitadoDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance53;
			appearance54.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance54.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdInvitadoDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance54;
			this.ugdInvitadoDetalle.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdInvitadoDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.ugdInvitadoDetalle.Location = new System.Drawing.Point(0, 40);
			this.ugdInvitadoDetalle.Name = "ugdInvitadoDetalle";
			this.ugdInvitadoDetalle.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ugdInvitadoDetalle.Size = new System.Drawing.Size(1104, 296);
			this.ugdInvitadoDetalle.TabIndex = 1061;
			this.ugdInvitadoDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugdInvitadoDetalle_InitializeLayout);
			// 
			// udsInvitadoDet
			// 
			ultraDataColumn27.DataType = typeof(System.DateTime);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(System.UInt32);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			this.udsInvitadoDet.Band.Columns.AddRange(new object[] {
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
																															 ultraDataColumn39});
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button2.CausesValidation = false;
			this.button2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button2.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.Location = new System.Drawing.Point(1092, -68);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(80, 55);
			this.button2.TabIndex = 1058;
			this.button2.Text = "Salir";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.button2, "Salir");
			// 
			// groupBox5
			// 
			this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox5.Controls.Add(this.label2);
			this.groupBox5.Controls.Add(this.ultraCombo1);
			this.groupBox5.Controls.Add(this.label12);
			this.groupBox5.Controls.Add(this.txtNumero);
			this.groupBox5.Controls.Add(this.cmbEstado);
			this.groupBox5.Controls.Add(this.lblEstado);
			this.groupBox5.Location = new System.Drawing.Point(-67, -76);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(1126, 76);
			this.groupBox5.TabIndex = 1056;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "Datos principales";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 24);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(51, 16);
			this.label2.TabIndex = 1038;
			this.label2.Text = "* Usuario";
			// 
			// ultraCombo1
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo1.Appearance = appearance55;
			this.ultraCombo1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo1.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn48.Header.VisiblePosition = 1;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 0;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn48,
																										 ultraGridColumn49});
			this.ultraCombo1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraCombo1.DisplayMember = "Nombre";
			this.ultraCombo1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo1.Location = new System.Drawing.Point(104, 24);
			this.ultraCombo1.Name = "ultraCombo1";
			this.ultraCombo1.ReadOnly = true;
			this.ultraCombo1.Size = new System.Drawing.Size(248, 21);
			this.ultraCombo1.TabIndex = 1042;
			this.ultraCombo1.ValueMember = "idPersonal";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(448, 24);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(39, 16);
			this.label12.TabIndex = 1051;
			this.label12.Text = "Buscar";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label12.Visible = false;
			// 
			// txtNumero
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance56;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(512, 24);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(120, 21);
			this.txtNumero.TabIndex = 1050;
			this.txtNumero.Visible = false;
			// 
			// cmbEstado
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance57;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn50.Header.VisiblePosition = 0;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 1;
			ultraGridColumn51.Width = 217;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn50,
																										 ultraGridColumn51});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(720, 24);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(80, 21);
			this.cmbEstado.TabIndex = 1055;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			this.cmbEstado.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Location = new System.Drawing.Point(648, 24);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(104, 23);
			this.lblEstado.TabIndex = 1054;
			this.lblEstado.Text = "* Estado:";
			this.lblEstado.Visible = false;
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
			// udsResponsable
			// 
			ultraDataColumn40.DataType = typeof(System.UInt32);
			this.udsResponsable.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn40,
																															 ultraDataColumn41});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.cmbEstadoFF);
			this.groupBox1.Controls.Add(this.label7);
			this.groupBox1.Controls.Add(this.cmbFFCabecera);
			this.groupBox1.Controls.Add(this.lblTipo);
			this.groupBox1.Controls.Add(this.txtBusqueda);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.cmbArea);
			this.groupBox1.Controls.Add(this.btnConsultar);
			this.groupBox1.Location = new System.Drawing.Point(13, 76);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(982, 100);
			this.groupBox1.TabIndex = 1059;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Datos principales";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// cmbEstadoFF
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstadoFF.Appearance = appearance58;
			this.cmbEstadoFF.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstadoFF.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 0;
			ultraGridColumn53.Width = 450;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn52,
																										 ultraGridColumn53});
			this.cmbEstadoFF.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbEstadoFF.DisplayMember = "";
			this.cmbEstadoFF.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstadoFF.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstadoFF.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstadoFF.Location = new System.Drawing.Point(408, 56);
			this.cmbEstadoFF.Name = "cmbEstadoFF";
			this.cmbEstadoFF.Size = new System.Drawing.Size(248, 21);
			this.cmbEstadoFF.TabIndex = 1048;
			this.cmbEstadoFF.ValueMember = "";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(344, 56);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(39, 16);
			this.label7.TabIndex = 1047;
			this.label7.Text = "Estado";
			// 
			// cmbFFCabecera
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFFCabecera.Appearance = appearance59;
			this.cmbFFCabecera.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFFCabecera.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn54.Header.VisiblePosition = 1;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn55.Width = 316;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn54,
																										 ultraGridColumn55});
			this.cmbFFCabecera.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbFFCabecera.DisplayMember = "";
			this.cmbFFCabecera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFFCabecera.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFFCabecera.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFFCabecera.Location = new System.Drawing.Point(64, 18);
			this.cmbFFCabecera.Name = "cmbFFCabecera";
			this.cmbFFCabecera.Size = new System.Drawing.Size(247, 21);
			this.cmbFFCabecera.TabIndex = 1046;
			this.cmbFFCabecera.ValueMember = "";
			// 
			// lblTipo
			// 
			this.lblTipo.AutoSize = true;
			this.lblTipo.Location = new System.Drawing.Point(16, 24);
			this.lblTipo.Name = "lblTipo";
			this.lblTipo.Size = new System.Drawing.Size(26, 16);
			this.lblTipo.TabIndex = 1045;
			this.lblTipo.Text = "Tipo";
			// 
			// txtBusqueda
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance60;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(408, 24);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(248, 21);
			this.txtBusqueda.TabIndex = 1044;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(344, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 1043;
			this.label4.Text = "Buscar";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 56);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(28, 16);
			this.label3.TabIndex = 1038;
			this.label3.Text = "Área";
			// 
			// cmbArea
			// 
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArea.Appearance = appearance61;
			this.cmbArea.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArea.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn56.Header.VisiblePosition = 1;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 0;
			ultraGridColumn57.Width = 316;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn56,
																										 ultraGridColumn57});
			this.cmbArea.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbArea.DisplayMember = "";
			this.cmbArea.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArea.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArea.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArea.Location = new System.Drawing.Point(64, 56);
			this.cmbArea.Name = "cmbArea";
			this.cmbArea.Size = new System.Drawing.Size(247, 21);
			this.cmbArea.TabIndex = 1042;
			this.cmbArea.ValueMember = "";
			// 
			// frmFamilyFriendsReporteria
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1140, 566);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox4);
			this.Name = "frmFamilyFriendsReporteria";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Reportería Family Friends";
			this.Load += new System.EventHandler(this.frmFamilyFriendsReporteria_Load);
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			this.groupBox4.ResumeLayout(false);
			this.tabControl1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgSeguimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsFamilyFriends_ReporteCon)).EndInit();
			this.tabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdInvitadoDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsInvitadoDet)).EndInit();
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadoFF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFFCabecera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArea)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region DataSets
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}		
		#endregion DataSets

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			FamilyFriendsCon();
			FamilyFriendsConDetalle();
		}

		private string ResolverArea()
		{
			if (cmbArea == null || cmbArea.Value == null)
				return string.Empty;
			switch (Convert.ToInt32(cmbArea.Value))
			{
				case 2: return "COMERCIAL";
				case 3: return "ADMINISTRATIVA";
				default: return string.Empty;
			}
		}

		private FamilyFriendsReporteFiltroModel CrearFiltroParam()
		{
			FamilyFriendsReporteFiltroModel param = new FamilyFriendsReporteFiltroModel();
			param.IdCabecera = new Guid(this.cmbFFCabecera.Value.ToString());
			param.Filtro = txtBusqueda.Text.Trim();
			param.Area = ResolverArea();
			param.idEstado = Convert.ToInt32(this.cmbEstadoFF.Value);
			return param;
		}

		private void FamilyFriendsCon()
		{
			this.label6.Text = "";

			try
			{
				MenuLatinium.MostrarModalProcesando();

				FamilyFriendsReporteService service = new FamilyFriendsReporteService(MenuLatinium.sconexionEmpresa);
				ArrayList lista = service.ObtenerReporte(CrearFiltroParam());

				DataTable dt = new DataTable();

				dt.Columns.Add("Row", typeof(int));
				dt.Columns.Add("Colaborador", typeof(string));
				dt.Columns.Add("EmpresaEmpleado", typeof(string));
				dt.Columns.Add("AreaNomina", typeof(string));
				dt.Columns.Add("DepartamentoNomina", typeof(string));
				dt.Columns.Add("CargoNomina", typeof(string));
				dt.Columns.Add("Identificacion", typeof(string));
				dt.Columns.Add("Nombre_Bodega", typeof(string));

				dt.Columns.Add("TotalInvitados", typeof(int));
				dt.Columns.Add("TotalFacturas", typeof(int));

				dt.Columns.Add("TotalMontoFacturas", typeof(decimal));
				dt.Columns.Add("TotalNotasCreditoCount", typeof(int));
				dt.Columns.Add("TotalNotasCredito", typeof(decimal));
				dt.Columns.Add("TotalNeto", typeof(decimal));

				dt.Columns.Add("TotalCosto", typeof(decimal));
				dt.Columns.Add("MargenUtilidad", typeof(decimal));

				foreach (FamilyFriendsReporteModel item in lista)
				{
					DataRow row = dt.NewRow();

					row["Colaborador"] = item.Colaborador;
					row["EmpresaEmpleado"] = item.EmpresaEmpleado;
					row["AreaNomina"] = item.AreaNomina;
					row["DepartamentoNomina"] = item.DepartamentoNomina;
					row["CargoNomina"] = item.CargoNomina;
					row["Identificacion"] = item.Identificacion;
					row["Nombre_Bodega"] = item.Nombre_Bodega;

					row["TotalInvitados"] = item.TotalInvitados;
					row["TotalFacturas"] = item.TotalFacturas;

					row["TotalMontoFacturas"] =
						item.TotalMontoFacturas;

					row["TotalNotasCreditoCount"] =
						item.TotalNotasCreditoCount;

					row["TotalNotasCredito"] =
						item.TotalNotasCredito;

					row["TotalNeto"] =
						item.TotalNeto;

					row["TotalCosto"] =
						item.TotalCosto;

					row["MargenUtilidad"] =
						item.MargenUtilidad;

					dt.Rows.Add(row);
				}

				this.udgSeguimiento.DataSource = dt;

				this.udgSeguimiento =
					Funcion.ContadorFilas(
					this.udgSeguimiento,
					"Row"
					);

				this.label6.Text =
					this.udgSeguimiento.Rows.Count +
					" REGISTROS ENCONTRADOS";
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					"Ocurrió un error al consultar el reporte: "
					+ ex.Message,
					"Error",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
			finally
			{
				MenuLatinium.OcultarModalProcesando();
			}
		}

		private void FamilyFriendsConDetalle()
		{
			this.label5.Text = "";

			try
			{
				MenuLatinium.MostrarModalProcesando();

				FamilyFriendsReporteService service = new FamilyFriendsReporteService(MenuLatinium.sconexionEmpresa);
				ArrayList lista = service.ObtenerReporteDetalle(CrearFiltroParam());

				DataTable dt = new DataTable();

				dt.Columns.Add("Row", typeof(int));
				dt.Columns.Add("Invitado", typeof(string));
				dt.Columns.Add("Cedula", typeof(string));
				dt.Columns.Add("Email", typeof(string));
				dt.Columns.Add("Celular", typeof(string));

				dt.Columns.Add("AreaNomina", typeof(string));
				dt.Columns.Add("PatrocinadorNombre", typeof(string));
				dt.Columns.Add("PatrocinadorCedula", typeof(string));

				dt.Columns.Add("TipoDocumento", typeof(string));
				dt.Columns.Add("NumeroDocumento", typeof(string));
				dt.Columns.Add("FechaDocumento", typeof(DateTime));

				// NUEVOS CAMPOS
				dt.Columns.Add("FF", typeof(string));
				dt.Columns.Add("Articulo", typeof(string));
				dt.Columns.Add("Codigo", typeof(string));
				dt.Columns.Add("Cantidad", typeof(decimal));
				dt.Columns.Add("DescuentoPorc", typeof(decimal));
				dt.Columns.Add("SubTotal", typeof(decimal));

				dt.Columns.Add("CostoTotalUltimo", typeof(decimal));
				dt.Columns.Add("MargenUtilidad", typeof(decimal));

				dt.Columns.Add("NombreBodega", typeof(string));
				dt.Columns.Add("FormaPago", typeof(string));

				foreach (FamilyFriendsInvitadoDetalleModel item in lista)
				{
					DataRow row = dt.NewRow();

					row["Invitado"] = item.Invitado;
					row["Cedula"] = item.Cedula;
					row["Email"] = item.Email;
					row["Celular"] = item.Celular;

					row["AreaNomina"] = item.AreaNomina;
					row["PatrocinadorNombre"] = item.PatrocinadorNombre;
					row["PatrocinadorCedula"] = item.PatrocinadorCedula;

					row["TipoDocumento"] = item.TipoDocumento;
					row["NumeroDocumento"] = item.NumeroDocumento;
					row["FechaDocumento"] = item.FechaDocumento;

					// NUEVOS VALORES
					row["FF"] = item.FF;
					row["Articulo"] = item.Articulo;
					row["Codigo"] = item.Codigo;
					row["Cantidad"] = item.Cantidad;
					row["DescuentoPorc"] = item.DescuentoPorc;
					row["SubTotal"] = item.SubTotal;

					row["CostoTotalUltimo"] = item.CostoTotalUltimo;

					// MARGEN CALCULADO EN UI
					row["MargenUtilidad"] =
						item.SubTotal - item.CostoTotalUltimo;

					row["NombreBodega"] = item.NombreBodega;
					row["FormaPago"] = item.FormaPago;

					dt.Rows.Add(row);
				}

				ugdInvitadoDetalle.DataSource = dt;

				ugdInvitadoDetalle.DisplayLayout.ViewStyleBand =
					Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;

				ugdInvitadoDetalle.DisplayLayout.GroupByBox.Hidden = false;

				Infragistics.Win.UltraWinGrid.UltraGridBand band =
					ugdInvitadoDetalle.DisplayLayout.Bands[0];

				// SIN AGRUPACIÓN INICIAL
				band.SortedColumns.Clear();

				band.Summaries.Clear();

				// SUBTOTAL
				Infragistics.Win.UltraWinGrid.SummarySettings sumSubTotal =
					band.Summaries.Add(
					"SumSubTotal",
					Infragistics.Win.UltraWinGrid.SummaryType.Sum,
					band.Columns["SubTotal"]
					);

				sumSubTotal.DisplayFormat = "{0:N2}";
				sumSubTotal.Appearance.TextHAlign =
					Infragistics.Win.HAlign.Right;

				// COSTO
				Infragistics.Win.UltraWinGrid.SummarySettings sumCosto =
					band.Summaries.Add(
					"SumCosto",
					Infragistics.Win.UltraWinGrid.SummaryType.Sum,
					band.Columns["CostoTotalUltimo"]
					);

				sumCosto.DisplayFormat = "{0:N2}";
				sumCosto.Appearance.TextHAlign =
					Infragistics.Win.HAlign.Right;

				// MARGEN
				Infragistics.Win.UltraWinGrid.SummarySettings sumMargen =
					band.Summaries.Add(
					"SumMargen",
					Infragistics.Win.UltraWinGrid.SummaryType.Sum,
					band.Columns["MargenUtilidad"]
					);

				sumMargen.DisplayFormat = "{0:N2}";
				sumMargen.Appearance.TextHAlign =
					Infragistics.Win.HAlign.Right;

				// CANTIDAD
				Infragistics.Win.UltraWinGrid.SummarySettings sumCantidad =
					band.Summaries.Add(
					"SumCantidad",
					Infragistics.Win.UltraWinGrid.SummaryType.Sum,
					band.Columns["Cantidad"]
					);

				sumCantidad.DisplayFormat = "{0:N0}";
				sumCantidad.Appearance.TextHAlign =
					Infragistics.Win.HAlign.Right;

				// DESCUENTO
				Infragistics.Win.UltraWinGrid.SummarySettings sumDescuento =
					band.Summaries.Add(
					"SumDescuentoPorc",
					Infragistics.Win.UltraWinGrid.SummaryType.Sum,
					band.Columns["DescuentoPorc"]
					);

				sumDescuento.DisplayFormat = "{0:N2}";
				sumDescuento.Appearance.TextHAlign =
					Infragistics.Win.HAlign.Right;

				ugdInvitadoDetalle.Rows.ExpandAll(true);

				this.ugdInvitadoDetalle =
					Funcion.ContadorFilas(
					this.ugdInvitadoDetalle,
					"Row"
					);

				this.label5.Text =
					dt.Rows.Count + " REGISTROS ENCONTRADOS";
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					"Ocurrió un error al consultar el reporte: "
					+ ex.Message,
					"Error",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
			finally
			{
				MenuLatinium.OcultarModalProcesando();
			}
		}

		private void CargarArea()
		{
			System.Data.DataTable dt = new System.Data.DataTable();

			// Columnas
			dt.Columns.Add("Id", typeof(int));
			dt.Columns.Add("Nombre", typeof(string));

			// Filas
			System.Data.DataRow r;

			r = dt.NewRow();
			r["Id"] = 1;
			r["Nombre"] = "TODO";
			dt.Rows.Add(r);

			r = dt.NewRow();
			r["Id"] = 2;
			r["Nombre"] = "COMERCIAL";
			dt.Rows.Add(r);

			r = dt.NewRow();
			r["Id"] = 3;
			r["Nombre"] = "ADMINISTRATIVA";
			dt.Rows.Add(r);

			// DataSource
			this.cmbArea.DataSource = dt;
			this.cmbArea.DisplayMember = "Nombre";
			this.cmbArea.ValueMember = "Id";

			// OCULTAR ID
			this.cmbArea.DisplayLayout.Bands[0]
				.Columns["Id"].Hidden = true;

			// DEFAULT
			this.cmbArea.Value = 1;
		}

		private void CargarCabecerasFF()
		{
			System.Data.DataTable dt = new System.Data.DataTable();

			FamilyFriendsReporteService service = null;
			ArrayList lista = null;

			// COLUMNAS
			dt.Columns.Add("IdCabecera", typeof(string));
			dt.Columns.Add("Titulo", typeof(string));

			try
			{
				service = new FamilyFriendsReporteService(
					MenuLatinium.sconexionEmpresa
					);

				// OBTENER CABECERAS
				lista = service.ObtenerCabeceras();

				// RECORRER LISTA (.NET 1.1)
				IEnumerator en = lista.GetEnumerator();

				while (en.MoveNext())
				{
					FamilyFriendsCabeceraModel item =
						(FamilyFriendsCabeceraModel)en.Current;

					System.Data.DataRow r = dt.NewRow();

					r["IdCabecera"] =
						item.IdCabecera.ToString();

					r["Titulo"] =
						item.Titulo;

					dt.Rows.Add(r);
				}

				// ASIGNAR DATASOURCE
				this.cmbFFCabecera.DataSource = dt;
				this.cmbFFCabecera.DisplayMember = "Titulo";
				this.cmbFFCabecera.ValueMember = "IdCabecera";

				// OCULTAR ID
				this.cmbFFCabecera.DisplayLayout.Bands[0]
					.Columns["IdCabecera"].Hidden = true;

				// ANCHO COLUMNA
				this.cmbFFCabecera.DisplayLayout.Bands[0]
					.Columns["Titulo"].Width = 450;

				// ANCHO DROPDOWN
				this.cmbFFCabecera.DropDownWidth = 450;

				// SELECCIONAR PRIMER ITEM
				if (this.cmbFFCabecera.Rows.Count > 0)
				{
					this.cmbFFCabecera.ActiveRow =
						this.cmbFFCabecera.Rows[0];
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					"Error al cargar cabeceras Family & Friends: "
					+ ex.Message,
					"Sistema",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
		}

		private void CargarEstadosFF()
		{
			System.Data.DataTable dt = new System.Data.DataTable();

			FamilyFriendsReporteService service = null;
			ArrayList lista = null;

			// COLUMNAS
			dt.Columns.Add("idEstado", typeof(int));
			dt.Columns.Add("NombreEstado", typeof(string));

			try
			{
				service = new FamilyFriendsReporteService(
					MenuLatinium.sconexionEmpresa
					);

				// OBTENER ESTADOS
				lista = service.ObtenerEstados();

				// RECORRER LISTA (.NET 1.1)
				IEnumerator en = lista.GetEnumerator();

				while (en.MoveNext())
				{
					FamilyFriendsEstadCombooModel item =
						(FamilyFriendsEstadCombooModel)en.Current;

					System.Data.DataRow r = dt.NewRow();

					r["idEstado"] = item.idEstado;
					r["NombreEstado"] = item.NombreEstado;

					dt.Rows.Add(r);
				}

				// ASIGNAR DATASOURCE
				this.cmbEstadoFF.DataSource = dt;
				this.cmbEstadoFF.DisplayMember = "NombreEstado";
				this.cmbEstadoFF.ValueMember = "idEstado";

				// OCULTAR ID
				this.cmbEstadoFF.DisplayLayout.Bands[0]
					.Columns["idEstado"].Hidden = true;

				// ANCHO COLUMNA
				this.cmbEstadoFF.DisplayLayout.Bands[0]
					.Columns["NombreEstado"].Width = 450;

				// ANCHO DROPDOWN
				this.cmbEstadoFF.DropDownWidth = 450;

				// SELECCIONAR PRIMER ITEM
				if (this.cmbEstadoFF.Rows.Count > 0)
				{
					this.cmbEstadoFF.ActiveRow =
						this.cmbEstadoFF.Rows[0];
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					"Error al cargar estados Family & Friends: "
					+ ex.Message,
					"Sistema",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
		}

		private void frmFamilyFriendsReporteria_Load(object sender, System.EventArgs e)
		{
			CargarArea();
			CargarCabecerasFF();
			CargarEstadosFF();
			int idUsuarioInicio = Funcion.iEjecucion(frmRequerimiento.sconexionInicio, string.Format("select top 1 idUsuario from Usuarios where NOMBRE = '{0}' Order By idUsuario", MenuLatinium.stUsuario),false);
			//Cargar el nombre de la empresa que se logeo
			string sempresa = Funcion.sEjecucion(frmRequerimiento.sconexionInicio, string.Format("Select Top 1 e.Empresa From Empresas e Where e.BaseDeDatos = '{0}'", MenuLatinium.stDirFacturacion.Trim()));
			//CargarResponsable
			this.cmbResponsable.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionInicio, "Exec RequerimientoResponsable");
			//Traer de la tabla de inicio los usuarios y el logeado
			try
			{
				//Seleccionar el personal que ingreso al sistema que es responsable
				this.cmbResponsable.Value = idUsuarioInicio;

				if((int)this.cmbResponsable.Value == 0)
				{
					MessageBox.Show(string.Format("Su usuario {0}, no se encuentra registrado como usuario de la empresa {1}", 
						MenuLatinium.stUsuario,
						sempresa
						), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}
			catch (Exception)
			{
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			try
			{
				// Mostrar el modal de "procesando"
				MenuLatinium.MostrarModalProcesando();
				FuncionesProcedimientos.ExportaExcel(this.udgSeguimiento);
			}
			catch (Exception ex)
			{
				// Manejar cualquier excepción que pueda ocurrir durante la consulta
				MessageBox.Show("Ocurrió un error durante la consulta: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				// Ocultar el modal de "procesando" una vez que se haya completado la consulta
				MenuLatinium.OcultarModalProcesando();
			}
		}

		private void udgSeguimiento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			e.Layout.Bands[0].Columns["TotalMontoFacturas"].Format = "N2";
			e.Layout.Bands[0].Columns["TotalNotasCredito"].Format = "N2";
			e.Layout.Bands[0].Columns["TotalNeto"].Format = "N2";
			e.Layout.Bands[0].Columns["TotalCosto"].Format = "N2";
			e.Layout.Bands[0].Columns["MargenUtilidad"].Format = "N2";
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			try
			{
				// Mostrar el modal de "procesando"
				MenuLatinium.MostrarModalProcesando();
				FuncionesProcedimientos.ExportaExcel(this.ugdInvitadoDetalle);
			}
			catch (Exception ex)
			{
				// Manejar cualquier excepción que pueda ocurrir durante la consulta
				MessageBox.Show("Ocurrió un error durante la consulta: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				// Ocultar el modal de "procesando" una vez que se haya completado la consulta
				MenuLatinium.OcultarModalProcesando();
			}
		}

		private void ugdInvitadoDetalle_InitializeLayout(
			object sender,
			Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			e.Layout.Bands[0].Columns["CostoTotalUltimo"].Format = "N2";
			e.Layout.Bands[0].Columns["MargenUtilidad"].Format = "N2";

			e.Layout.Bands[0].Columns["Cantidad"].Format = "N0";
			e.Layout.Bands[0].Columns["DescuentoPorc"].Format = "N2";
			e.Layout.Bands[0].Columns["SubTotal"].Format = "N2";
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}
	}
}
