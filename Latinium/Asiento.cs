using System;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using LibContabilidad.DataObjects;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Asientos.
	/// </summary>
	public class Asiento : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsAsiento;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCuenta;
		private C1.Data.C1DataView cdvDetAsiento;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private C1.Data.C1DataView cdvTipoAsiento;
		private DevExpress.XtraEditors.TextEdit txtIdAsiento;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsSeteoC;
		private C1.Data.C1DataSet cdsSeguridad;
		private DevExpress.XtraEditors.DateEdit dtFechaIngreso;
		private System.Windows.Forms.Label label11;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataView cdvSubProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private DevExpress.XtraEditors.SimpleButton btCheque;
		private System.Windows.Forms.ImageList imageList;
		private DevExpress.XtraEditors.SimpleButton btPlanCuenta;
		private C1.Data.C1DataSet cdsMoneda;
		private C1.Data.C1DataView cdvMoneda;
		private DevExpress.XtraEditors.SimpleButton btSaldo;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private C1.Data.C1DataView cdvAsiento;
		private C1.Data.C1DataView cdvCuentaCod;
		private Infragistics.Win.UltraWinGrid.UltraGrid grAsiento;
		private System.Windows.Forms.Label lblAnulado;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmdProyecto;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalD;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalH;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSubProy;
		private System.Windows.Forms.HelpProvider helpProvider1;
		private System.Windows.Forms.Label lblCheque;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbDescrpcion;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCodCuenta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtValDolar;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaReg;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMoneda;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private DevExpress.XtraEditors.SimpleButton btNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCiudad;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsCiudad;
		private DevExpress.XtraEditors.SimpleButton btDigitalizacion;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbNoDeducible;
		private Infragistics.Win.Misc.UltraButton btAuditoria;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCheque;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private System.ComponentModel.IContainer components;

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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Asiento));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_AsientoDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbDescrpcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmdProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto", -1, "cmbSubProy");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DetIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaConf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIng");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto", -1, "cmdProyecto");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaComp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCostoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProyectoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNoDeducible", -1, "cmbNoDeducible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("jmid");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCod", 0, "cmbCodCuenta");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debe", 6, true, "_AsientoDetalle", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debe", 6, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Haber", 7, true, "_AsientoDetalle", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Haber", 7, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CajaChica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroCosto");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("AsientoTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MonedaTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoMoneda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCiudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNoDeducible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cdvDetAsiento = new C1.Data.C1DataView();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.cdvTipoAsiento = new C1.Data.C1DataView();
			this.txtIdAsiento = new DevExpress.XtraEditors.TextEdit();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.dtFechaIngreso = new DevExpress.XtraEditors.DateEdit();
			this.cdvMoneda = new C1.Data.C1DataView();
			this.cdsMoneda = new C1.Data.C1DataSet();
			this.lblCheque = new System.Windows.Forms.Label();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			this.label11 = new System.Windows.Forms.Label();
			this.btCheque = new DevExpress.XtraEditors.SimpleButton();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.btPlanCuenta = new DevExpress.XtraEditors.SimpleButton();
			this.btSaldo = new DevExpress.XtraEditors.SimpleButton();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.cdvAsiento = new C1.Data.C1DataView();
			this.cdvCuentaCod = new C1.Data.C1DataView();
			this.grAsiento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbDescrpcion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCodCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.lblAnulado = new System.Windows.Forms.Label();
			this.cmdProyecto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtTotalD = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalH = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cmbSubProy = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.helpProvider1 = new System.Windows.Forms.HelpProvider();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtValDolar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.cmbFechaReg = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSubProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cmbMoneda = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCiudad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsCiudad = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btDigitalizacion = new DevExpress.XtraEditors.SimpleButton();
			this.btNumero = new DevExpress.XtraEditors.SimpleButton();
			this.cmbNoDeducible = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btAuditoria = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label5 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvDetAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMoneda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMoneda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaCod)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDescrpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalH)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProy)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValDolar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaReg)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMoneda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNoDeducible)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsAsiento";
			this.cdsAsiento.FillOnRequest = false;
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			this.cdsAsiento.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsAsiento_PositionChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(152, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 3;
			this.label1.Text = "Código:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(152, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 16);
			this.label2.TabIndex = 4;
			this.label2.Text = "Número:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(328, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(46, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "Descrip:";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(328, 40);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(31, 16);
			this.label4.TabIndex = 8;
			this.label4.Text = "Nota:";
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.RowFilter = "Grupo = 0";
			this.cdvCuenta.Sort = "Descripcion";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// cdvDetAsiento
			// 
			this.cdvDetAsiento.BindingContextCtrl = this;
			this.cdvDetAsiento.DataSet = this.cdsAsiento;
			this.cdvDetAsiento.TableName = "";
			this.cdvDetAsiento.TableViewName = "AsientoDetalle";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 43);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(44, 16);
			this.label6.TabIndex = 36;
			this.label6.Text = "Numero";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(336, 43);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(32, 16);
			this.label7.TabIndex = 35;
			this.label7.Text = "Nota:";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.ForeColor = System.Drawing.Color.Black;
			this.label8.Location = new System.Drawing.Point(336, 11);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(61, 16);
			this.label8.TabIndex = 34;
			this.label8.Text = "Descripción";
			// 
			// cdvTipoAsiento
			// 
			this.cdvTipoAsiento.BindingContextCtrl = this;
			this.cdvTipoAsiento.DataSet = this.cdsAsiento;
			this.cdvTipoAsiento.RowFilter = "Grupo = 0";
			this.cdvTipoAsiento.TableName = "";
			this.cdvTipoAsiento.TableViewName = "AsientoTipo";
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsAsiento, "_Asiento.idAsiento"));
			this.txtIdAsiento.EditValue = ((object)(resources.GetObject("txtIdAsiento.EditValue")));
			this.txtIdAsiento.Location = new System.Drawing.Point(520, 456);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.Size = new System.Drawing.Size(40, 20);
			this.txtIdAsiento.TabIndex = 38;
			this.txtIdAsiento.ToolTip = "Código de asiento";
			this.txtIdAsiento.Visible = false;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			this.cdsSeteoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoC_BeforeFill);
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			this.cdsSeguridad.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeguridad_BeforeFill);
			// 
			// dtFechaIngreso
			// 
			this.dtFechaIngreso.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsAsiento, "_Asiento.FechaIngreso"));
			this.dtFechaIngreso.EditValue = ((object)(resources.GetObject("dtFechaIngreso.EditValue")));
			this.dtFechaIngreso.Location = new System.Drawing.Point(464, 456);
			this.dtFechaIngreso.Name = "dtFechaIngreso";
			// 
			// dtFechaIngreso.Properties
			// 
			this.dtFechaIngreso.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaIngreso.Size = new System.Drawing.Size(56, 20);
			this.dtFechaIngreso.TabIndex = 39;
			this.dtFechaIngreso.Visible = false;
			// 
			// cdvMoneda
			// 
			this.cdvMoneda.BindingContextCtrl = this;
			this.cdvMoneda.DataSet = this.cdsMoneda;
			this.cdvMoneda.TableName = "";
			this.cdvMoneda.TableViewName = "MonedaTipo";
			// 
			// cdsMoneda
			// 
			this.cdsMoneda.BindingContextCtrl = this;
			this.cdsMoneda.DataLibrary = "LibContabilidad";
			this.cdsMoneda.DataLibraryUrl = "";
			this.cdsMoneda.DataSetDef = "dsMoneda";
			this.cdsMoneda.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsMoneda.Name = "cdsMoneda";
			this.cdsMoneda.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsMoneda.SchemaDef = null;
			this.cdsMoneda.BeforeFill += new C1.Data.FillEventHandler(this.cdsMoneda_BeforeFill);
			// 
			// lblCheque
			// 
			this.lblCheque.AutoSize = true;
			this.lblCheque.Location = new System.Drawing.Point(224, 75);
			this.lblCheque.Name = "lblCheque";
			this.lblCheque.Size = new System.Drawing.Size(42, 16);
			this.lblCheque.TabIndex = 44;
			this.lblCheque.Text = "Cheque";
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible =1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			this.cdsProyecto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.RowFilter = "Visible = 1";
			this.cdvSubProyecto.Sort = "Nombre";
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(624, 104);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(47, 16);
			this.label11.TabIndex = 48;
			this.label11.Text = "Proyecto";
			this.label11.Visible = false;
			this.label11.MouseDown += new System.Windows.Forms.MouseEventHandler(this.label11_MouseDown);
			// 
			// btCheque
			// 
			this.btCheque.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCheque.ImageIndex = 0;
			this.btCheque.ImageList = this.imageList;
			this.btCheque.Location = new System.Drawing.Point(312, 416);
			this.btCheque.Name = "btCheque";
			this.btCheque.Size = new System.Drawing.Size(24, 24);
			this.btCheque.TabIndex = 13;
			this.btCheque.Text = "Cheque";
			this.btCheque.ToolTip = "Impresion de Cheque";
			this.btCheque.Click += new System.EventHandler(this.btCheque_Click);
			// 
			// imageList
			// 
			this.imageList.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btPlanCuenta
			// 
			this.btPlanCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btPlanCuenta.ImageIndex = 2;
			this.btPlanCuenta.ImageList = this.imageList;
			this.btPlanCuenta.Location = new System.Drawing.Point(336, 416);
			this.btPlanCuenta.Name = "btPlanCuenta";
			this.btPlanCuenta.Size = new System.Drawing.Size(24, 24);
			this.btPlanCuenta.TabIndex = 15;
			this.btPlanCuenta.Text = "Cheque";
			this.btPlanCuenta.ToolTip = "Permiter aumentar o modificar Plan de Cuentas";
			this.btPlanCuenta.Click += new System.EventHandler(this.btPlanCuenta_Click);
			this.btPlanCuenta.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btPlanCuenta_MouseDown);
			// 
			// btSaldo
			// 
			this.btSaldo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSaldo.ImageIndex = 3;
			this.btSaldo.ImageList = this.imageList;
			this.btSaldo.Location = new System.Drawing.Point(360, 416);
			this.btSaldo.Name = "btSaldo";
			this.btSaldo.Size = new System.Drawing.Size(24, 24);
			this.btSaldo.TabIndex = 16;
			this.btSaldo.Text = "Saldo";
			this.btSaldo.ToolTip = "Saldo de Cuenta seleccionado / Presupuesto (Click secundario)";
			this.btSaldo.Click += new System.EventHandler(this.btSaldo_Click);
			this.btSaldo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btSaldo_MouseDown);
			// 
			// cdsCuentaCorr
			// 
			this.cdsCuentaCorr.BindingContextCtrl = this;
			this.cdsCuentaCorr.DataLibrary = "LibContabilidad";
			this.cdsCuentaCorr.DataLibraryUrl = "";
			this.cdsCuentaCorr.DataSetDef = "dsCuentaCorriente";
			this.cdsCuentaCorr.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaCorr.Name = "cdsCuentaCorr";
			this.cdsCuentaCorr.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaCorr.SchemaDef = null;
			this.cdsCuentaCorr.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentaCorr_BeforeFill);
			// 
			// cdvAsiento
			// 
			this.cdvAsiento.BindingContextCtrl = this;
			this.cdvAsiento.DataSet = this.cdsAsiento;
			this.cdvAsiento.Sort = "idAsiento";
			this.cdvAsiento.TableName = "";
			this.cdvAsiento.TableViewName = "Asiento";
			// 
			// cdvCuentaCod
			// 
			this.cdvCuentaCod.BindingContextCtrl = this;
			this.cdvCuentaCod.DataSet = this.cdsCuenta;
			this.cdvCuentaCod.RowFilter = "Grupo =0";
			this.cdvCuentaCod.Sort = "Codigo";
			this.cdvCuentaCod.TableName = "";
			this.cdvCuentaCod.TableViewName = "Cuenta";
			// 
			// grAsiento
			// 
			this.grAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grAsiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.grAsiento.DataMember = "_Asiento.Asiento - AsientoDetalle";
			this.grAsiento.DataSource = this.cdsAsiento;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grAsiento.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 19;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 18;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 21;
			ultraGridColumn3.Hidden = true;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Header.Caption = "Descripción";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn4.Width = 254;
			ultraGridColumn5.Header.Caption = "C. Costo";
			ultraGridColumn5.Header.VisiblePosition = 20;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn6.Header.Caption = "Proyecto";
			ultraGridColumn6.Header.VisiblePosition = 12;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeBoth;
			ultraGridColumn7.MaskInput = "-n,nnn,nnn,nnn.nn";
			ultraGridColumn7.PromptChar = ' ';
			ultraGridColumn7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridColumn7.Width = 97;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.MaskInput = "-n,nnn,nnn,nnn.nn";
			ultraGridColumn8.PromptChar = ' ';
			ultraGridColumn8.Width = 103;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 162;
			ultraGridColumn10.Header.VisiblePosition = 13;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 15;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 16;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 17;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Hidden = true;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridColumn15.CellAppearance = appearance6;
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 8;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 210;
			ultraGridColumn20.Header.Caption = "CentroCosto";
			ultraGridColumn20.Header.VisiblePosition = 9;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.Caption = "Proyecto";
			ultraGridColumn21.Header.VisiblePosition = 11;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 22;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn23.Header.VisiblePosition = 23;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 24;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 25;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 26;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 27;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 28;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 29;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 30;
			ultraGridColumn30.Hidden = true;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridColumn31.CellAppearance = appearance7;
			ultraGridColumn31.Header.Caption = "Código";
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn31.Width = 128;
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
																										 ultraGridColumn31});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance8;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance9;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "";
			this.grAsiento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.grAsiento.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.grAsiento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grAsiento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grAsiento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.grAsiento.DisplayLayout.Override.CardAreaAppearance = appearance11;
			this.grAsiento.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 10F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grAsiento.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grAsiento.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grAsiento.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grAsiento.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.grAsiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grAsiento.Location = new System.Drawing.Point(8, 136);
			this.grAsiento.Name = "grAsiento";
			this.grAsiento.Size = new System.Drawing.Size(888, 272);
			this.grAsiento.TabIndex = 8;
			this.grAsiento.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.grAsiento.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.grAsiento_AfterRowInsert);
			this.grAsiento.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.grAsiento_BeforeRowsDeleted);
			this.grAsiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grAsiento_KeyDown);
			this.grAsiento.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.grAsiento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grAsiento_InitializeLayout);
			// 
			// cmbDescrpcion
			// 
			this.cmbDescrpcion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDescrpcion.DataSource = this.cdvCuenta;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn33.Header.VisiblePosition = 4;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 3;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 2;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 0;
			ultraGridColumn36.Width = 283;
			ultraGridColumn37.Header.VisiblePosition = 5;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			this.cmbDescrpcion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbDescrpcion.DisplayMember = "Descripcion";
			this.cmbDescrpcion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDescrpcion.Location = new System.Drawing.Point(16, 248);
			this.cmbDescrpcion.Name = "cmbDescrpcion";
			this.cmbDescrpcion.Size = new System.Drawing.Size(376, 86);
			this.cmbDescrpcion.TabIndex = 55;
			this.cmbDescrpcion.Text = "cmDescripcion";
			this.cmbDescrpcion.ValueMember = "idCuenta";
			this.cmbDescrpcion.Visible = false;
			// 
			// cmbCodCuenta
			// 
			this.cmbCodCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCodCuenta.DataSource = this.cdvCuentaCod;
			ultraGridColumn38.Header.VisiblePosition = 2;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 4;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 3;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.Caption = "Código";
			ultraGridColumn41.Header.VisiblePosition = 0;
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.Width = 288;
			ultraGridColumn43.Header.VisiblePosition = 5;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43});
			this.cmbCodCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCodCuenta.DisplayMember = "CodRapido";
			this.cmbCodCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodCuenta.Location = new System.Drawing.Point(320, 328);
			this.cmbCodCuenta.Name = "cmbCodCuenta";
			this.cmbCodCuenta.Size = new System.Drawing.Size(400, 86);
			this.cmbCodCuenta.TabIndex = 56;
			this.cmbCodCuenta.Text = "cmbCodCuenta";
			this.cmbCodCuenta.ValueMember = "idCuenta";
			this.cmbCodCuenta.Visible = false;
			// 
			// lblAnulado
			// 
			this.lblAnulado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblAnulado.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.lblAnulado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblAnulado.ForeColor = System.Drawing.Color.Red;
			this.lblAnulado.Location = new System.Drawing.Point(560, 416);
			this.lblAnulado.Name = "lblAnulado";
			this.lblAnulado.Size = new System.Drawing.Size(120, 25);
			this.lblAnulado.TabIndex = 59;
			this.lblAnulado.Text = "ANULADO";
			this.lblAnulado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblAnulado.Visible = false;
			// 
			// cmdProyecto
			// 
			this.cmdProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdProyecto.DataSource = this.cdvProyecto;
			ultraGridColumn44.Header.VisiblePosition = 2;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 6;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Width = 154;
			ultraGridColumn47.Header.VisiblePosition = 3;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 7;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 4;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.Caption = "Centro Costo";
			ultraGridColumn50.Header.VisiblePosition = 1;
			ultraGridColumn51.Header.VisiblePosition = 5;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 8;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 9;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 10;
			ultraGridColumn55.Header.VisiblePosition = 11;
			ultraGridColumn56.Header.VisiblePosition = 12;
			ultraGridBand4.Columns.AddRange(new object[] {
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
			this.cmdProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmdProyecto.DisplayMember = "Nombre";
			this.cmdProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdProyecto.Location = new System.Drawing.Point(64, 328);
			this.cmdProyecto.Name = "cmdProyecto";
			this.cmdProyecto.Size = new System.Drawing.Size(256, 86);
			this.cmdProyecto.TabIndex = 60;
			this.cmdProyecto.ValueMember = "idProyecto";
			this.cmdProyecto.Visible = false;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.CausesValidation = false;
			this.barraDato1.DataMember = "_Asiento";
			this.barraDato1.DataNombreId = "idAsiento";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsAsiento;
			this.barraDato1.DataTabla = "Asiento";
			this.barraDato1.DataTablaHija = "AsientoDetalle";
			this.barraDato1.Location = new System.Drawing.Point(8, 415);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 0;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtTotalD
			// 
			this.txtTotalD.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance16.FontData.BoldAsString = "True";
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalD.Appearance = appearance16;
			this.txtTotalD.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.Total"));
			this.txtTotalD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalD.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotalD.FormatString = "#,##0.00";
			this.txtTotalD.Location = new System.Drawing.Point(688, 416);
			this.txtTotalD.Name = "txtTotalD";
			this.txtTotalD.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalD.PromptChar = ' ';
			this.txtTotalD.ReadOnly = true;
			this.txtTotalD.Size = new System.Drawing.Size(104, 24);
			this.txtTotalD.TabIndex = 62;
			// 
			// txtTotalH
			// 
			this.txtTotalH.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalH.Appearance = appearance17;
			this.txtTotalH.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.TotalH"));
			this.txtTotalH.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalH.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotalH.FormatString = "#,##0.00";
			this.txtTotalH.Location = new System.Drawing.Point(792, 416);
			this.txtTotalH.Name = "txtTotalH";
			this.txtTotalH.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalH.PromptChar = ' ';
			this.txtTotalH.ReadOnly = true;
			this.txtTotalH.Size = new System.Drawing.Size(104, 24);
			this.txtTotalH.TabIndex = 63;
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// cmbSubProy
			// 
			this.cmbSubProy.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubProy.DataSource = this.cdvSubProyecto;
			ultraGridColumn57.Header.VisiblePosition = 3;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 2;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 0;
			ultraGridColumn59.Width = 161;
			ultraGridColumn60.Header.VisiblePosition = 4;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 6;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 5;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 1;
			ultraGridColumn63.Width = 89;
			ultraGridColumn64.Header.VisiblePosition = 7;
			ultraGridColumn64.Hidden = true;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64});
			this.cmbSubProy.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbSubProy.DisplayMember = "Nombre";
			this.cmbSubProy.Location = new System.Drawing.Point(392, 248);
			this.cmbSubProy.Name = "cmbSubProy";
			this.cmbSubProy.Size = new System.Drawing.Size(272, 104);
			this.cmbSubProy.TabIndex = 65;
			this.cmbSubProy.ValueMember = "idSubProyecto";
			this.cmbSubProy.Visible = false;
			// 
			// helpProvider1
			// 
			this.helpProvider1.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance18;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.Notas"));
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtNotas.Location = new System.Drawing.Point(400, 40);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(496, 88);
			this.txtNotas.TabIndex = 7;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance19;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.Descripcion"));
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtDescripcion.Location = new System.Drawing.Point(400, 8);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(320, 22);
			this.txtDescripcion.TabIndex = 3;
			this.txtDescripcion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescripcion_KeyDown);
			this.txtDescripcion.ValueChanged += new System.EventHandler(this.txtDescripcion_ValueChanged);
			// 
			// txtCodigo
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance20;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.CodAsiento"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtCodigo.Location = new System.Drawing.Point(104, 40);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(104, 22);
			this.txtCodigo.TabIndex = 4;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.ValueChanged += new System.EventHandler(this.txtCodigo_ValueChanged);
			// 
			// txtNumero
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance21;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtNumero.Location = new System.Drawing.Point(104, 72);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 5;
			this.txtNumero.MouseUp += new System.Windows.Forms.MouseEventHandler(this.txtNumero_MouseUp);
			// 
			// txtValDolar
			// 
			this.txtValDolar.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.ValDolar"));
			this.txtValDolar.Location = new System.Drawing.Point(560, 456);
			this.txtValDolar.Name = "txtValDolar";
			this.txtValDolar.Size = new System.Drawing.Size(48, 22);
			this.txtValDolar.TabIndex = 70;
			this.txtValDolar.Visible = false;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblUsuario.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsAsiento, "_Asiento.Usuario"));
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblUsuario.ForeColor = System.Drawing.Color.Blue;
			this.lblUsuario.Location = new System.Drawing.Point(440, 418);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(15, 20);
			this.lblUsuario.TabIndex = 71;
			this.lblUsuario.Text = "U";
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbFechaReg
			// 
			this.cmbFechaReg.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaReg.Appearance = appearance22;
			this.cmbFechaReg.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.FechaRef"));
			dateButton1.Caption = "Today";
			this.cmbFechaReg.DateButtons.Add(dateButton1);
			this.cmbFechaReg.Format = "dd/MMM/yyyy";
			this.cmbFechaReg.Location = new System.Drawing.Point(800, 104);
			this.cmbFechaReg.Name = "cmbFechaReg";
			this.cmbFechaReg.NonAutoSizeHeight = 24;
			this.cmbFechaReg.Size = new System.Drawing.Size(96, 21);
			this.cmbFechaReg.TabIndex = 10;
			this.toolTip1.SetToolTip(this.cmbFechaReg, "Fecha de Referencia");
			this.cmbFechaReg.Value = new System.DateTime(2006, 12, 7, 0, 0, 0, 0);
			this.cmbFechaReg.Visible = false;
			// 
			// cmbProyecto
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance23;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.idProyecto"));
			this.cmbProyecto.DataSource = this.ultraDataSource1;
			ultraGridColumn65.Header.VisiblePosition = 0;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 1;
			ultraGridColumn66.Width = 300;
			ultraGridColumn67.Header.VisiblePosition = 2;
			ultraGridColumn67.Hidden = true;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(104, 104);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(288, 21);
			this.cmbProyecto.TabIndex = 6;
			this.cmbProyecto.ValueMember = "idProyecto";
			this.cmbProyecto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbProyecto_KeyDown);
			this.cmbProyecto.Validating += new System.ComponentModel.CancelEventHandler(this.cmbProyecto_Validating);
			this.cmbProyecto.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbProyecto_MouseUp);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// cmbSubProyecto
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubProyecto.Appearance = appearance24;
			this.cmbSubProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.idSubProyecto"));
			this.cmbSubProyecto.DataSource = this.cdvSubProyecto;
			ultraGridColumn68.Header.VisiblePosition = 3;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 2;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 0;
			ultraGridColumn70.Width = 243;
			ultraGridColumn71.Header.VisiblePosition = 4;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 6;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 5;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 1;
			ultraGridColumn75.Header.VisiblePosition = 7;
			ultraGridColumn75.Hidden = true;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn68,
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73,
																										 ultraGridColumn74,
																										 ultraGridColumn75});
			this.cmbSubProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbSubProyecto.DisplayMember = "Nombre";
			this.cmbSubProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubProyecto.Location = new System.Drawing.Point(672, 104);
			this.cmbSubProyecto.Name = "cmbSubProyecto";
			this.cmbSubProyecto.Size = new System.Drawing.Size(128, 21);
			this.cmbSubProyecto.TabIndex = 6;
			this.cmbSubProyecto.ValueMember = "idSubProyecto";
			this.cmbSubProyecto.Visible = false;
			this.cmbSubProyecto.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbSubProyecto_MouseUp);
			// 
			// cmbTipo
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance25;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.idTipoAsiento"));
			this.cmbTipo.DataSource = this.cdvTipoAsiento;
			ultraGridColumn76.Header.VisiblePosition = 5;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 0;
			ultraGridColumn78.Header.VisiblePosition = 7;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 1;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 2;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 6;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 4;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 3;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 8;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79,
																										 ultraGridColumn80,
																										 ultraGridColumn81,
																										 ultraGridColumn82,
																										 ultraGridColumn83,
																										 ultraGridColumn84});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(8, 8);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(152, 21);
			this.cmbTipo.TabIndex = 1;
			this.toolTip1.SetToolTip(this.cmbTipo, "Tipo de Asiento");
			this.cmbTipo.ValueMember = "idTipoAsiento";
			this.cmbTipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipo_KeyDown);
			this.cmbTipo.Validating += new System.ComponentModel.CancelEventHandler(this.cmbTipo_Validating);
			this.cmbTipo.ValueChanged += new System.EventHandler(this.cmbTipo_ValueChanged);
			this.cmbTipo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipo_InitializeLayout);
			this.cmbTipo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbTipo_MouseUp);
			// 
			// dtFecha
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance26;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.Fecha"));
			dateButton2.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton2);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(216, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.TabIndex = 2;
			this.toolTip1.SetToolTip(this.dtFecha, "Fecha de Contabilizacion");
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			// 
			// cmbMoneda
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMoneda.Appearance = appearance27;
			this.cmbMoneda.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMoneda.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMoneda.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.idTipoMoneda"));
			this.cmbMoneda.DataSource = this.cdvMoneda;
			ultraGridColumn85.Header.VisiblePosition = 2;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 0;
			ultraGridColumn87.Header.VisiblePosition = 1;
			ultraGridColumn87.Hidden = true;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn85,
																										 ultraGridColumn86,
																										 ultraGridColumn87});
			this.cmbMoneda.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbMoneda.DisplayMember = "Nombre";
			this.cmbMoneda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMoneda.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMoneda.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMoneda.Location = new System.Drawing.Point(512, 104);
			this.cmbMoneda.Name = "cmbMoneda";
			this.cmbMoneda.Size = new System.Drawing.Size(104, 21);
			this.cmbMoneda.TabIndex = 7;
			this.toolTip1.SetToolTip(this.cmbMoneda, "Moneda de asiento");
			this.cmbMoneda.ValueMember = "idTipoMoneda";
			this.cmbMoneda.Visible = false;
			this.cmbMoneda.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbMoneda_MouseDown);
			// 
			// cmbCiudad
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCiudad.Appearance = appearance28;
			this.cmbCiudad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCiudad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCiudad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.idCiudad"));
			this.cmbCiudad.DataSource = this.udsCiudad;
			ultraGridBand10.AddButtonCaption = "MonedaTipo";
			ultraGridColumn88.Header.VisiblePosition = 1;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 0;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn88,
																											ultraGridColumn89});
			this.cmbCiudad.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbCiudad.DisplayMember = "Ciudad";
			this.cmbCiudad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCiudad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCiudad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCiudad.Location = new System.Drawing.Point(400, 104);
			this.cmbCiudad.Name = "cmbCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(104, 21);
			this.cmbCiudad.TabIndex = 74;
			this.toolTip1.SetToolTip(this.cmbCiudad, "Ciudad del cheque");
			this.cmbCiudad.ValueMember = "idCiudad";
			this.cmbCiudad.Visible = false;
			// 
			// udsCiudad
			// 
			this.udsCiudad.Band.Columns.AddRange(new object[] {
																													ultraDataColumn4,
																													ultraDataColumn5});
			// 
			// btDigitalizacion
			// 
			this.btDigitalizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btDigitalizacion.ImageIndex = 1;
			this.btDigitalizacion.ImageList = this.imageList;
			this.btDigitalizacion.Location = new System.Drawing.Point(408, 416);
			this.btDigitalizacion.Name = "btDigitalizacion";
			this.btDigitalizacion.Size = new System.Drawing.Size(24, 24);
			this.btDigitalizacion.TabIndex = 75;
			this.toolTip1.SetToolTip(this.btDigitalizacion, "Digitalización");
			this.btDigitalizacion.Visible = false;
			this.btDigitalizacion.Click += new System.EventHandler(this.btDigitalizacion_Click);
			this.btDigitalizacion.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btDigitalizacion_MouseDown);
			// 
			// btNumero
			// 
			this.btNumero.Location = new System.Drawing.Point(8, 71);
			this.btNumero.Name = "btNumero";
			this.btNumero.Size = new System.Drawing.Size(80, 25);
			this.btNumero.TabIndex = 73;
			this.btNumero.Text = "Referencia";
			this.btNumero.Click += new System.EventHandler(this.btNumero_Click);
			// 
			// cmbNoDeducible
			// 
			this.cmbNoDeducible.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn90.Header.VisiblePosition = 0;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 1;
			ultraGridColumn91.Width = 149;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn90,
																											ultraGridColumn91});
			this.cmbNoDeducible.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbNoDeducible.DisplayMember = "Detalle";
			this.cmbNoDeducible.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNoDeducible.Location = new System.Drawing.Point(720, 312);
			this.cmbNoDeducible.Name = "cmbNoDeducible";
			this.cmbNoDeducible.Size = new System.Drawing.Size(164, 104);
			this.cmbNoDeducible.TabIndex = 76;
			this.cmbNoDeducible.ValueMember = "idNoDeducible";
			this.cmbNoDeducible.Visible = false;
			// 
			// btAuditoria
			// 
			this.btAuditoria.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance29.Image = ((object)(resources.GetObject("appearance29.Image")));
			this.btAuditoria.Appearance = appearance29;
			this.btAuditoria.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAuditoria.Location = new System.Drawing.Point(384, 416);
			this.btAuditoria.Name = "btAuditoria";
			this.btAuditoria.Size = new System.Drawing.Size(24, 24);
			this.btAuditoria.TabIndex = 123;
			this.btAuditoria.Click += new System.EventHandler(this.btAuditoria_Click);
			this.btAuditoria.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAuditoria_MouseUp);
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
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(176, 11);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(34, 16);
			this.label5.TabIndex = 124;
			this.label5.Text = "Fecha";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(8, 107);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(85, 16);
			this.label9.TabIndex = 125;
			this.label9.Text = "Centro de Costo";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCheque
			// 
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			appearance30.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCheque.Appearance = appearance30;
			this.txtCheque.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.txtCheque.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCheque.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.Cheque"));
			this.txtCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtCheque.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtCheque.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.txtCheque.Location = new System.Drawing.Point(280, 72);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.Size = new System.Drawing.Size(112, 22);
			this.txtCheque.TabIndex = 8;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.ForeColor = System.Drawing.Color.Black;
			this.label10.Location = new System.Drawing.Point(736, 11);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(36, 16);
			this.label10.TabIndex = 127;
			this.label10.Text = "R.U.C.";
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor1.Appearance = appearance31;
			this.ultraTextEditor1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsiento, "_Asiento.RUC"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.ultraTextEditor1.Location = new System.Drawing.Point(784, 8);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(112, 22);
			this.ultraTextEditor1.TabIndex = 126;
			// 
			// Asiento
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(904, 446);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.lblCheque);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.btAuditoria);
			this.Controls.Add(this.cmbNoDeducible);
			this.Controls.Add(this.btDigitalizacion);
			this.Controls.Add(this.cmbCiudad);
			this.Controls.Add(this.btNumero);
			this.Controls.Add(this.cmbMoneda);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbSubProyecto);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.cmbFechaReg);
			this.Controls.Add(this.txtCheque);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.cmbSubProy);
			this.Controls.Add(this.txtTotalH);
			this.Controls.Add(this.txtTotalD);
			this.Controls.Add(this.cmdProyecto);
			this.Controls.Add(this.lblAnulado);
			this.Controls.Add(this.cmbCodCuenta);
			this.Controls.Add(this.cmbDescrpcion);
			this.Controls.Add(this.grAsiento);
			this.Controls.Add(this.btSaldo);
			this.Controls.Add(this.btPlanCuenta);
			this.Controls.Add(this.btCheque);
			this.Controls.Add(this.dtFechaIngreso);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtValDolar);
			this.helpProvider1.SetHelpKeyword(this, "Asiento1.Htm");
			this.helpProvider1.SetHelpNavigator(this, System.Windows.Forms.HelpNavigator.Topic);
			this.helpProvider1.SetHelpString(this, "");
			this.Name = "Asiento";
			this.helpProvider1.SetShowHelp(this, true);
			this.Text = "Asientos";
			this.Resize += new System.EventHandler(this.Asiento_Resize);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Asiento_Closing);
			this.Load += new System.EventHandler(this.Asientos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvDetAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMoneda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMoneda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaCod)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDescrpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalH)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProy)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValDolar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaReg)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMoneda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNoDeducible)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables Globales
		private int IdAsiento = 0;
		private C1DataRow drSeteoF;
		SeteoCRow_tableView drSeteoC;
		C1.Data.C1DataRow drSeteo;
		public Acceso miAcceso;
		C1.Data.C1DataRow drAsi;
		private int iCuentaCuentasC = 0;
		int [] ListaCuentasC = new int[200];
		private bool bActualiza = true;
		string controlW="";
		string controlW1="";
		string NuevoCodigo="";
	
		#endregion Variables Globales
		#region Pantalla Asiento
		public Asiento()
		{
			InitializeComponent();
		}

		public Asiento(int idAsi)
		{
			IdAsiento = idAsi;
			InitializeComponent();
		}
		private void UnloadMe()
		{
			this.Close();
		}
		private void Asientos_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("44",cdsSeguridad))
			{
				MessageBox.Show("No puede ingresar a Asientos", "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			#region Auditoria
			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 6, 'Asiento'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
			#endregion Auditoria
			#region Ordenacion de grilla
			if (Funcion.bEjecutaSQL(cdsCuenta, "Exec SeteoGExiste 'ACOG'")) 				
				grAsiento.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			#endregion Ordenacion de grilla

			#region Seguridad
			miAcceso =new Acceso("0402", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			if (!Funcion.Permiso("187", cdsSeguridad)) txtCheque.ReadOnly = true;
			if (!Funcion.Permiso("188", cdsSeguridad)) txtNumero.ReadOnly = true;
			#endregion Seguridad
			Personalizaciones();
			LlenaCombos();
			#region Valores Iniciales
			txtValDolar.Width = 0;
			dtFechaIngreso.Width = 0;
			txtIdAsiento.Width = 0;
			drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];
			C1DataRow dr = cdsSeteoC.TableViews["SeteoC"].Rows[0];
			drSeteoC = SeteoCRow_tableView.Obj(dr);
			LlenaListaCuentasC();
			#endregion Valores Iniciales
			#region Version academica
			if (MenuLatinium.bAcademica)
			{
				string stCuenta = "Select Count(*) From Asiento";
				int iCuentaTotal = Funcion.iEscalarSQL(cdsAsiento, stCuenta, true);
				if (iCuentaTotal > 200) barraDato1.sbtNuevo.Enabled = false;
			}
			#endregion Version academica
			#region Visibilidad

			// Conciliacion Tributaria
			if (Funcion.bEjecutaSQL(cdsAsiento, "Exec SeteoGExiste 'AVCT'")) 				
			{
				cmbNoDeducible.DataSource = Funcion.dvProcedimiento(cdsSeteoC, "Select idNiif_TipoDTDP As idNoDeducible, Concepto AS Detalle from dbo.Niif_TipoDTDP");
				grAsiento.DisplayLayout.Bands[0].Columns["idNoDeducible"].Hidden = false;										
			}

			if (Funcion.bEjecutaSQL(cdsAsiento, "Exec SeteoGExiste 'ADIGDEC'")) 				
			{
				btDigitalizacion.Visible=true;
			}

			bool bVerCiudad = Funcion.bEscalarSQL(cdsAsiento, "Exec SeteoGExiste 'AVC'", true);
			if (bVerCiudad)
			{
				cmbCiudad.DataSource = Funcion.dvProcedimiento(cdsAsiento, "Exec ClienteCiudadLista");
				cmbCiudad.Visible = true;
			}
			barraDato1.BarraMovimiento = miAcceso.Buscar;
			if (!miAcceso.Anular && !miAcceso.Eliminar)
				barraDato1.sbtBorrar.Enabled = false;
			barraDato1.sbtEditar.Enabled = miAcceso.Editar;
			barraDato1.sbtNuevo.Enabled = miAcceso.Nuevo;
			barraDato1.sbtImprimir.Enabled = miAcceso.Imprimir;
			btAuditoria.Visible = miAcceso.Auditoria;
			#endregion Visibilidad
			#region Busqueda y Habilitacion
			if (IdAsiento > 0)
				BuscaAsiento(IdAsiento);
			else
				barraDato1.PosUltima();

			if (drAsi == null || drAsi["idAsiento"] == DBNull.Value)
			{
				barraDato1.PosUltima();
				MessageBox.Show("No existen Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			else
				IdAsiento = (int) drAsi["idAsiento"];//drAsi["idAsiento"];
			barraDato1.HabilitaControles(false);
			ModificaGrilla(false);
			#endregion Busqueda y Habilitacion
		}
		private void Personalizaciones()
		{
			#region Proyectos por Centro de Costo BOTROSA
			if (Funcion.bEjecutaSQL(cdsAsiento, "Exec SeteoGExiste 'FTRPBP'")) 				
			{								
				grAsiento.DisplayLayout.Bands[0].Columns["idProyecto"].Hidden = true;										
				grAsiento.DisplayLayout.Bands[0].Columns["idSubProyecto"].Hidden = true;				
				grAsiento.DisplayLayout.Bands[0].Columns["ProyectoBR"].Hidden = false;										
				grAsiento.DisplayLayout.Bands[0].Columns["CentroCostoBr"].Hidden = false;
			}
			#endregion Botrosa
		}

		private void LlenaListaCuentasC()
		{
			SqlDataReader dr = Funcion.rEscalarSQL(cdsAsiento, "Select idCuenta From CuentaCorriente", true);
			int i = 0;
			while (dr.Read())
			{
				ListaCuentasC[i++] = dr.GetInt32(0);
			}
			dr.Close();
			iCuentaCuentasC = i;
		}
		private void LlenaCombos()
		{		
			#region Proyecto Altropico
			string stExec = "Exec ProyectoCarga " + 1;
			cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
			#endregion Proyecto Altropico
			//Select idOrigenAsiento, Descripcion From OrigenAsientos Order By Descripcion
		}

		private void ModificaGrilla(bool bModifica)
		{
			grAsiento.Enabled = true;
			if (bModifica)
			{
				grAsiento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				grAsiento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				grAsiento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				grAsiento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void Asiento_Resize(object sender, System.EventArgs e)
		{
			grAsiento.DisplayLayout.AutoFitColumns = true;
		}

		private void Asiento_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (barraDato1.bEditar == true)
			{
				MessageBox.Show("No puede salir mientras edita un Asiento", "Información");
				e.Cancel = true;
			}
			string stAudita = "Exec AuditaCrear 53, 7, 'Asiento'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}


		#endregion Pantalla Asiento
		#region Barra Estandard
		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (frmBusquedaDeAsientos Buscar = new frmBusquedaDeAsientos())
			{
				if (Buscar.ShowDialog() == DialogResult.OK)
				{
					try
					{						
						int id = int.Parse(Buscar.uGridAsientos.ActiveRow.Cells["idAsiento"].Value.ToString());
						BuscaAsiento(id);
					}
					catch{}
				}
			}
		}

		public void BuscaAsiento(int idAsi)
		{
			if (idAsi > 0) 
			{
				barraDato1.IdRegistro = idAsi;
				barraDato1.ProximoId(5);
			}
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a creación de Asientos", "Confirmación", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

//			int idTipoAsi = 1;
//			if (cmbTipo.Value != System.DBNull.Value) idTipoAsi = (int) cmbTipo.Value;
			barraDato1.CrearRegistro();

			drAsi["Fecha"] = DateTime.Today;
			drAsi["FechaRef"] = DateTime.Today;
//			drAsi["idTipoAsiento"] = 1;
//			drAsi["idSucursal"] = (int) drSeteo["NumEmpresa"];
			this.cmbTipo.Focus();
			ModificaGrilla(true);
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea copiar este registro?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			int id = CopiarAsiento();
			string stExec = "Exec AsientoCheque " +id.ToString();
			Funcion.EjecutaSQL(cdsAsiento, stExec);

			BuscaAsiento(id);
		}

		private int CopiarAsiento()
		{
			C1.Data.C1DataRow drC = cdsAsiento.TableViews["Asiento"].AddNew();
			drC["idProyecto"] = drAsi["idProyecto"];
			drC["idSubProyecto"] = drAsi["idSubProyecto"];
			drC["idTipoAsiento"] = drAsi["idTipoAsiento"];
			drC["idTipoMoneda"] = drAsi["idTipoMoneda"];
			drC["Fecha"] = DateTime.Today;
			drC["Total"] = drAsi["Total"];
			drC["TotalH"] = drAsi["Total"];
			string stNumero = Contabilidad.NumeroAsiento(cdsAsiento, (int) cmbTipo.Value);
			if (stNumero.Trim().Length > 0) drC["CodAsiento"] = stNumero;
			drC["Numero"] = drAsi["Numero"];
			drC["ValDolar"] = drAsi["ValDolar"];
			drC["Descripcion"] = drAsi["Descripcion"];
			drC["Notas"] = drAsi["Notas"];

			cdsAsiento.Update();

			// Copia de lineas
			foreach(C1.Data.C1DataRow dr in drAsi.GetChildRows())
			{
				C1.Data.C1DataRow drDC = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
				drDC["idAsiento"] = drC["idAsiento"];
				drDC["idCuenta"] = dr["idCuenta"];
				drDC["Debe"] = dr["Debe"];
				drDC["Haber"] = dr["Haber"];
				drDC["Detalle"] = dr["Detalle"];
				drDC["CentroCosto"] = dr["CentroCosto"];
				drDC["idProyecto"] = dr["idProyecto"];
				drDC["idSubProyecto"] = dr["idSubProyecto"];
			}
			cdsAsiento.Update();
			return (int) drC["idAsiento"];
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			barraDato1.Focus();

//			if (this.cmbProyecto.ActiveRow == null)
//			{
//				MessageBox.Show("Seleccione un centro de costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.cmbProyecto.Focus();
//				return;
//			}
			
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Contabilidad"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Grabar este Asiento, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				return;
			}

			double lnDiferencia = (double) txtTotalD.Value - (double) txtTotalH.Value;
			if (Math.Abs(lnDiferencia) > .0001)
			{
				MessageBox.Show("ASIENTO DESCUADRADO DIFERENCIA DE " + lnDiferencia.ToString("#,##0.00"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!Verifica()) return;

			if ((int)this.cmbTipo.Value == 3)
			{
				MessageBox.Show("No puede crear Egreso en este Modulo, Solo Desde Pagos Autorizados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			bool bNuevo = false;

			if (barraDato1.bNuevo) 
			{
				string stNumero = Contabilidad.NumeroAsiento(cdsAsiento, (int) cmbTipo.Value);
				if (stNumero.Trim().Length > 0) 
				{
					drAsi["CodAsiento"] = stNumero;
					txtCodigo.Text = stNumero;
				}
				bNuevo = true;
			}
			else
			{
				drAsi["FechaEditado"] = DateTime.Now;
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
			{
				if ((dr.Cells["idProyecto"].Value == System.DBNull.Value || (int)dr.Cells["idProyecto"].Value == 0 ||
					dr.Cells["idSubProyecto"].Value == System.DBNull.Value || (int)dr.Cells["idSubProyecto"].Value == 0) &&
					dr.Cells["idCuenta"].Value != System.DBNull.Value)
				{
					int idCuenta = (int) dr.Cells["idCuenta"].Value;
					int iCuenta = Funcion.iEscalarSQL(cdsAsiento, "Select Count(*) From CuentaObliga where idCuenta = " + idCuenta);
					if (iCuenta > 0)
					{
						string stCuenta = "Select Codigo + ' ' + Descripcion From Cuenta Where idCuenta = " + idCuenta;
						stCuenta = Funcion.sEscalarSQL(cdsAsiento, stCuenta);
						string stMensaje = string.Format("Cuenta {0} Necesita Proyecto y Centro de Costo", stCuenta);
						MessageBox.Show(stMensaje, "Información");
						return;
					}
				}
			}
			if (!VerificaCodigo()) return;
			txtTotalD.Refresh();

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
			{
				if ((double)dr.Cells["Debe"].Value == 0 && (double)dr.Cells["Haber"].Value == 0)
				{
					MessageBox.Show("No puede Grabar En Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.grAsiento.ActiveRow = dr;
					this.grAsiento.ActiveRow.Selected = true;
					return;
				}
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
			{
				if (dr.Cells["idProyecto"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione el Centro de Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.grAsiento.ActiveRow = dr;
					this.grAsiento.ActiveRow.Selected = true;
					return;
				}

				if ((int)dr.Cells["idProyecto"].Value == 0)
				{
					MessageBox.Show("Seleccione el Centro de Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.grAsiento.ActiveRow = dr;
					this.grAsiento.ActiveRow.Selected = true;
					return;
				}
			}
		
			barraDato1.GrabaRegistro();
			if (bNuevo)
			{
				string stExec = "Exec AsientoCheque " +drAsi["idAsiento"].ToString();
				Funcion.EjecutaSQL(cdsAsiento, stExec);
			}
			barraDato1.IdRegistro = (int) drAsi["idAsiento"];
			barraDato1.ProximoId(5); 
			ModificaGrilla(false);
		}

		private bool VerificaOrigen()
		{
			string stEditarAutomatico = Funcion.sEscalarSQL(cdsAsiento, "Exec AsientoOrigenAutomatico " + drAsi["idAsiento"]);
			if (stEditarAutomatico.Length > 3)
			{
				int idPos = stEditarAutomatico.IndexOf(' ');
				idPos = stEditarAutomatico.IndexOf(' ', idPos + 1);
				MessageBox.Show(stEditarAutomatico.Substring(idPos), "Información", 
					MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return false;
			}
			return true;
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			drAsi = cdsAsiento.TableViews["Asiento"].Rows[0];
			if (drAsi == null)
			{
				MessageBox.Show("No existe un registro para editar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (!VerificaOrigen()) return;

			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Contabilidad"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Editar este Asiento, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				return;
			}

			if ((bool) drAsi["Bloquear"])
			{
				MessageBox.Show("No puede Editar Asiento.\nAsiento Bloqueado por el Administrador del Sistema",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a editar",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (miAcceso.EditarSoloDia)
			{
				DateTime dtFechaHoy = Funcion.dtEscalarSQL(cdsAsiento, "Select GetDate()");
				DateTime dtFechaAsi = (DateTime) dtFecha.Value;
				if (dtFechaAsi.ToShortDateString() != dtFechaHoy.ToShortDateString())
				{
					MessageBox.Show("Solo puede editar asientos de fecha " +
						dtFechaHoy.ToString("dd/MMM/yyyy"), 
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			if (miAcceso.EditarDiaIngreso)
			{
				DateTime dtFechaAsi = Funcion.dtEscalarSQL(cdsAsiento, "Select GetDate()");
        DateTime dtFechaIng = (DateTime) dtFechaIngreso.EditValue;
				if (dtFechaAsi.ToShortDateString() != dtFechaIng.ToShortDateString())
				{	
					MessageBox.Show("Asiento Ingresado el " +
						dtFechaIng.ToString("dd/MMM/yyyy") + 
            ".\nSolo puede editar Asientos Ingresados el día de Hoy",
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}

			barraDato1.EditarRegistro();
			ModificaGrilla(true);

			if (!miAcceso.CambiarFecha) dtFecha.Enabled = false;
			if (!miAcceso.CambiarNumeracion) 
			{
				txtCodigo.Enabled = false;
				cmbTipo.Enabled = false;
			}
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			cdsAsiento.RejectChanges();
			barraDato1.DeshacerRegistro();
			BorraErrores();
			ModificaGrilla(false);
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Imprimir(0);
		}

		private void Imprimir(int iTipoImp)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (this.cmbTipo.Value == System.DBNull.Value)
			{
				MessageBox.Show("Defina Tipo de Asiento");
				return;
			}
			Cursor = Cursors.WaitCursor;
			
			int IdTipo = (int)this.cmbTipo.Value;
			string strReporte = "";
			double fValorCheque = 0;
			string strFiltro = "{Asiento.idAsiento} = " + drAsi["idAsiento"].ToString();

			switch (IdTipo)
			{
				case 1:					
				case 2:
					strReporte = "Asientos.rpt";					
					break;
				case 3:
					strReporte = "Egreso.rpt";
					break;
				case 4:
					strReporte = "NotaDebito.rpt";
					fValorCheque = ValorCheque();
					break;
				case 5:
					strReporte = "NotaCredito.rpt";
					fValorCheque = ValorCheque();
					break;
				case 6:
					strReporte = "NotaTransferencia.rpt";
					fValorCheque = ValorCheque();
					break;
				default:
					strReporte = "NotaTransferencia.rpt";
					fValorCheque = ValorCheque();
					break;
			}

			if ((int) this.cmbTipo.Value == 3)
			{
				int idTipo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec TipoAutorizacionAsiento {0}", (int)this.drAsi["idAsiento"]));

				if (idTipo > 0)
				{
					#region Si viene del modulo de pagos
					ParameterFields paramFields = new ParameterFields ();
				
					ParameterField pfidAsiento = new ParameterField ();
					ParameterDiscreteValue ValIdAsiento = new ParameterDiscreteValue ();
					pfidAsiento.ParameterFieldName = "@idAsiento";
					ValIdAsiento.Value =  (int)this.drAsi["idAsiento"];
					pfidAsiento.CurrentValues.Add (ValIdAsiento);
					paramFields.Add (pfidAsiento);
				
					ParameterField pfidTipo = new ParameterField ();
					ParameterDiscreteValue ValidTipo = new ParameterDiscreteValue ();
					pfidTipo.ParameterFieldName = "@idTipo";
					ValidTipo.Value = idTipo;
					pfidTipo.CurrentValues.Add (ValidTipo);
					paramFields.Add (pfidTipo);
				
					ParameterField pfidAsientoDet = new ParameterField ();
					ParameterDiscreteValue ValIdAsientoDet = new ParameterDiscreteValue ();
					pfidAsientoDet.ParameterFieldName = "@idAsientoDet";
					ValIdAsientoDet.Value =  (int)this.drAsi["idAsiento"];
					pfidAsientoDet.CurrentValues.Add (ValIdAsientoDet);
					paramFields.Add (pfidAsientoDet);
				
					Reporte miRepor = new Reporte("ComprobanteDeEgreso.rpt", "");
					miRepor.MdiParent = this.MdiParent;
					miRepor.Titulo("Egreso");
					miRepor.crVista.ParameterFieldInfo = paramFields;
					miRepor.Show();	

					Cursor = Cursors.Default;
					return;
					#endregion Si viene del modulo de pagos
				}
				else
				{
					#region Si no viene del modulo de pagos
					bool bImprimeIngreso = false;
					string stReportePersonalizado = Funcion.sEscalarSQL(cdsAsiento, "Select IsNull(Reporte, '') From AsientoTipo Where idTipoAsiento = " + IdTipo.ToString() );
					if (stReportePersonalizado.Length > 0) strReporte = stReportePersonalizado;
									
					ParameterFields paramFields = new ParameterFields ();
					ParameterField paramField = new ParameterField ();

					ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
					paramField.ParameterFieldName = "ValorCheque";
					discreteVal.Value = fValorCheque;
					paramField.CurrentValues.Add (discreteVal);
					paramFields.Add (paramField);

					string[] stCuentaBanco = CuentaBanco();
					if (drSeteoC.ImprimeIngreso == true) bImprimeIngreso = true;
					
					ParameterField paramField1 = new ParameterField ();
					paramField1.ParameterFieldName = "BancoNombre";
					ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
					discreteVal1.Value = stCuentaBanco[0];
					paramField1.CurrentValues.Add (discreteVal1);
					paramFields.Add (paramField1);

					ParameterField paramField2 = new ParameterField ();
					paramField2.ParameterFieldName = "CuentaNombre";
					ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
					discreteVal2.Value = stCuentaBanco[1];
					paramField2.CurrentValues.Add (discreteVal2);
					paramFields.Add (paramField2);
					
					Reporte miRep = new Reporte(strReporte, strFiltro);

					if (iTipoImp == 0)
					{
						miRep.crVista.ParameterFieldInfo = paramFields;
						miRep.MdiParent = MdiParent;
						miRep.Show();
					}
					else
					{
						miRep.oRpt.SetParameterValue("ValorCheque", fValorCheque);
						miRep.oRpt.SetParameterValue("BancoNombre", stCuentaBanco[0]);
						miRep.oRpt.SetParameterValue("CuentaNombre", stCuentaBanco[1]);
						miRep.ImprimeRep(1);
					}
					Cursor = Cursors.Default;
					return;
					#endregion Si viene del modulo de pagos
				}
			}
			if ((int) this.cmbTipo.Value > 3)
			{
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramField = new ParameterField ();

				ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
				paramField.ParameterFieldName = "ValorCheque";
				discreteVal.Value = fValorCheque;
				paramField.CurrentValues.Add (discreteVal);
				paramFields.Add (paramField);			

				Reporte miRep = new Reporte(strReporte, strFiltro);

				string[] stCuentaBanco = CuentaBanco();

				if (iTipoImp == 0)
				{
					miRep.crVista.ParameterFieldInfo = paramFields;
					miRep.MdiParent = MdiParent;
					miRep.Show();
				}
				else
				{
					miRep.oRpt.SetParameterValue("ValorCheque", fValorCheque);
					miRep.oRpt.SetParameterValue("BancoNombre", stCuentaBanco[0]);
					miRep.oRpt.SetParameterValue("CuentaNombre", stCuentaBanco[1]);
					miRep.ImprimeRep(1);
				}
			}
			else
			{
				Reporte miReporte = new Reporte(strReporte, strFiltro);

				if (iTipoImp == 0)
				{
					miReporte.MdiParent = MdiParent;
					miReporte.Show();
				}
				else miReporte.ImprimeRep(1);
			}
			
			#region Codigo Anterior
			/*string strReporte = "";
			double fValorCheque = 0;
			string strFiltro = "{Asiento.idAsiento} = " + drAsi["idAsiento"].ToString();
			int IdTipo = (int) cmbTipo.Value;

			switch (IdTipo)
			{
				case 1:
					strReporte = "Diario.rpt";
					break;
				case 2:
					strReporte = "Ingreso.rpt";
					fValorCheque = ValorCheque();
					break;
				case 3:
					strReporte = "Egreso.rpt";
					fValorCheque = ValorCheque();
					break;
				case 4:
					strReporte = "NotaDebito.rpt";
					fValorCheque = ValorCheque();
					break;
				case 5:
					strReporte = "NotaCredito.rpt";
					fValorCheque = ValorCheque();
					break;
				case 6:
					strReporte = "NotaTransferencia.rpt";
					fValorCheque = ValorCheque();
					break;
				default:
					strReporte = "NotaTransferencia.rpt";
					fValorCheque = ValorCheque();
					break;
			}
			string stReportePersonalizado = Funcion.sEscalarSQL(cdsAsiento, "Select IsNull(Reporte, '') From AsientoTipo Where idTipoAsiento = " + IdTipo.ToString() );
			if (stReportePersonalizado.Length > 0) strReporte = stReportePersonalizado;

			//Caso de Ingresos y Egresos envia valor de cheque
			if ((int) cmbTipo.Value != 1)
			{
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramField = new ParameterField ();

				ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
				paramField.ParameterFieldName = "ValorCheque";
				discreteVal.Value = fValorCheque;
				paramField.CurrentValues.Add (discreteVal);
				paramFields.Add (paramField);

				// Egreso
				string[] stCuentaBanco = CuentaBanco();
				bool bImprimeIngreso = false;
				if (drSeteoC.ImprimeIngreso == true) bImprimeIngreso = true;
				if ((int) cmbTipo.Value == 3 || bImprimeIngreso)
				{
					ParameterField paramField1 = new ParameterField ();
					paramField1.ParameterFieldName = "BancoNombre";
					ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
					discreteVal1.Value = stCuentaBanco[0];
					paramField1.CurrentValues.Add (discreteVal1);
					paramFields.Add (paramField1);

					ParameterField paramField2 = new ParameterField ();
					paramField2.ParameterFieldName = "CuentaNombre";
					ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
					discreteVal2.Value = stCuentaBanco[1];
					paramField2.CurrentValues.Add (discreteVal2);
					paramFields.Add (paramField2);
				}

				Reporte miRep = new Reporte(strReporte, strFiltro);
				if (iTipoImp == 0)
				{
					miRep.crVista.ParameterFieldInfo = paramFields;
					miRep.MdiParent = MdiParent;
					miRep.Show();
				}
				else
				{
					miRep.oRpt.SetParameterValue("ValorCheque", fValorCheque);
					miRep.oRpt.SetParameterValue("BancoNombre", stCuentaBanco[0]);
					miRep.oRpt.SetParameterValue("CuentaNombre", stCuentaBanco[1]);
					miRep.ImprimeRep(1);
				}
			}
			else
			{
				Reporte miReporte = new Reporte(strReporte, strFiltro);
				if (iTipoImp == 0)
				{
					miReporte.MdiParent = MdiParent;
					miReporte.Show();
				}
				else
				{
					miReporte.ImprimeRep(1);
				}
			}*/
			#endregion Codigo Anterior

			Cursor = Cursors.Default;
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
			Imprimir(1);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			drAsi = cdsAsiento.TableViews["Asiento"].Rows[0];

			if (drAsi == null)
			{
				MessageBox.Show("No existe un registro para borrar", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Contabilidad"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Anular/Borrar este Asiento, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				return;
			}
			if ((bool) drAsi["Bloquear"])
			{
				MessageBox.Show("No puede Borrar Asiento.\nAsiento Bloqueado por el Administrador del Sistema", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if ((bool) drAsi["Borrar"] == false)
			{
				if (!miAcceso.Anular) 
				{
					MessageBox.Show("No tiene acceso a Anular Asientos", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				if (Funcion.bEscalarSQL(cdsSeteoC, string.Format("Exec ValidaAsientoTB {0}", (int)drAsi["idAsiento"]), true))
				{
					MessageBox.Show("No puede Anular: \n\n\n Este registro fue creado desde Transferencias Bancarias, \n\n debe Anularlo desde ese Modulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
	
				if (DialogResult.No == MessageBox.Show("¿Desea Anular el Asiento?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)) return;

				lblAnulado.Visible = true;
				drAsi["Borrar"] = true;
				cdsAsiento.Update();

				if ((int)drAsi["idTipoAsiento"] == 1)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grAsiento.Rows.All)
					{
						if (dr.Cells["idFactura"].Value != System.DBNull.Value)
						{
							if ((int)dr.Cells["idFactura"].Value > 0)
							{
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoFacturas {0}, 1, 9", (int)dr.Cells["idFactura"].Value));
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set Bloquear = 0, idAsientoLoteSolidario = 0 Where idCompra = {0}", (int)dr.Cells["idFactura"].Value));
							}
						}
					}
				}

				if ((int)drAsi["idTipoAsiento"] == 3) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Estado = 1 Where idAsiento = {0}", (int)drAsi["idAsiento"]));
			}
			else
			{
				if (!miAcceso.Eliminar || (drSeteo["EliminaAsiento"] == System.DBNull.Value || (bool) drSeteo["EliminaAsiento"] == false))
				{
					MessageBox.Show("No tiene acceso a Eliminar Asientos. (Seguridades o Arch. Prop. Generales)",
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (DialogResult.No == MessageBox.Show("¿Desea Borrar el Asiento?\n" + "Este proceso NO puede ser Revertido.", "Confirmación", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2) ) return;
				string stBorrar = "Delete AsientoDetalle Where idAsiento = " + drAsi["idAsiento"].ToString();
				Funcion.EjecutaSQL(cdsAsiento, stBorrar, true);
				barraDato1.BorraRegistro();
			}
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Recuperar)
			{
				MessageBox.Show("No tiene acceso a Recuperación de Asientos Anulados",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}
			if (!(bool) drAsi["Borrar"])
			{
				MessageBox.Show("Acción solo es permitida cuando el registro esta anulado",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}
			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Contabilidad"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Restaurar este Asiento, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea Restaurar el Asiento Anulado?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question ))
				return;
			drAsi = cdsAsiento.TableViews["Asiento"].Rows[0];
			drAsi["Borrar"] = false;
			lblAnulado.Visible = false;
			cdsAsiento.Update();
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drAsi = cdsAsiento.TableViews["Asiento"].Rows[0];
				if ((bool)drAsi["Borrar"] == true) lblAnulado.Visible = true;
				else lblAnulado.Visible = false;
			}
			catch{};
		}

		private void BorraErrores()
		{
			errorProvider.SetError(grAsiento, "");
			errorProvider.SetError(txtDescripcion, "");
			errorProvider.SetError(dtFecha, "");
			errorProvider.SetError(txtCheque, "");
		}

	
		#endregion Barra Estandard
		#region Barra Extendida
		private void btCheque_Click(object sender, System.EventArgs e)
		{
			if (!miAcceso.ImprimirCheque)
			{
				MessageBox.Show("No tiene acceso a impresión de cheques", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			double fValorCheque = 0;
			int iTipo = (int) cmbTipo.Value;
			if (iTipo == 3 || iTipo == 2)
			{
				fValorCheque = ValorCheque();
			}
			else
			{
				MessageBox.Show("Solo puede imprimir cheques en ingresos y egresos", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			// Definicion de Parametros
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();

			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "ValorCheque";
			discreteVal.Value = fValorCheque;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
			// Fin de definicion de parametros


			string strFiltro = "{Asiento.idAsiento} = " + drAsi["idAsiento"];
			string stCheque = "Cheque.rpt";
			string stSelect = "SELECT Reporte FROM CuentaCorriente Inner Join AsientoDetalle "
				+ "On AsientoDetalle.idCuenta= CuentaCorriente.idCuenta "
				+ "Where AsientoDetalle.idAsiento = " + txtIdAsiento.Text;
			if (iTipo == 2) stSelect += " And Debe > 0 ";
			else stSelect += " And Haber > 0 ";

			string stCheque1 = "";
			try
			{
				stCheque1 = Funcion.sEscalarSQL(cdsAsiento, stSelect, false);
				if (stCheque1.Length > 0) 
				{
					stCheque = stCheque1;
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show("No tiene una cuenta de la lista de cuentas corrientes: " + ex.Message, "Información");
			}

			Reporte miRep = new Reporte(stCheque, strFiltro);
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.MdiParent = MdiParent;
			miRep.Show();
		}

		private void btPlanCuenta_Click(object sender, System.EventArgs e)
		{
			using (PlanCuentas miPlan = new PlanCuentas())
			{
				miPlan.ShowDialog();
			}
			cdsCuenta.Clear();
			cdsCuenta.Fill();
		}

		private void btSaldo_Click(object sender, System.EventArgs e)
		{
			if (grAsiento.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Cuenta que desea ver el saldo");
				return;
			}
			if (grAsiento.ActiveRow.Cells["idCuenta"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione Cuenta que desea ver el saldo");
				return;
			}

			string stSelect = "SELECT IsNull(SUM(IsNull(Debe, 0) " +
				"- IsNull(Haber, 0)), 0) FROM AsientoDetalle WHERE idCuenta = " +
				grAsiento.ActiveRow.Cells["idCuenta"].Value.ToString() +
				" AND Borrar = 0";
			double fSaldo= Funcion.dEscalarSQL(cdsAsiento, stSelect);
			stSelect = "SELECT IsNull(SUM(Isnull(Debe, 0) - " +
				"IsNull(Haber, 0)),0) FROM PresupuestoDetalle WHERE idCuenta = " +
				grAsiento.ActiveRow.Cells["idCuenta"].Value.ToString() +
				" AND Borrar = 0";
			double fSaldoP = Funcion.dEscalarSQL(cdsAsiento, stSelect);
			double fDisponible = fSaldoP - fSaldo;
			stSelect = "Saldo Contable: " + fSaldo.ToString("#,##0.00")
				+ "\nSaldo Presupuestario: " + fSaldoP.ToString("#,##0.00")
				+ "\nSaldo Disponible: " + fDisponible.ToString("#,##0.00");
			MessageBox.Show(stSelect, "Saldo de Cuenta",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btAuditoria_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(23, (int) drAsi["idAsiento"], true);
			Cursor = Cursors.Default;
		}

		private void btAuditoria_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(23, (int) drAsi["idAsiento"], false);
			Cursor = Cursors.Default;		
		}

		private void btPlanCuenta_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			string strFiltro = "{Asiento.idAsiento} = " 
				+ drAsi["idAsiento"].ToString();
		
//1 Ajuste
//2 Ingreso
//3 Egreso
			string Report="";
			if ((int)drAsi["idtipoasiento"]==1 )
					Report="AsientoGrupoA";
			else if ((int)drAsi["idtipoasiento"]==2 )
					Report="AsientoGrupoI";
			else if ((int)drAsi["idtipoasiento"]==3 )
				Report="AsientoGrupo";
				using (Reporte miReporte = new Reporte(Report, strFiltro))
			{
				miReporte.ShowDialog();
			}
		}

		private void btSaldo_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (grAsiento.ActiveRow == null) return;
			int IdDetAsi = (int) grAsiento.ActiveRow.Cells["idDetAsiento"].Value;
			using (AsientoDist miDist = new AsientoDist(IdDetAsi))
			{
				try
				{
					miDist.ShowDialog();
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message);
				}
				barraDato1.ProximoId(5);
			}
		}

		private void btNumero_Click(object sender, System.EventArgs e)
		{
			int idAsi = int.Parse(txtIdAsiento.Text);
			//Busca en Factura
			string stSelect = "";

			// Busca en factura directa
			stSelect = "Select Count(*) From Compra Where idAsiento = " + idAsi.ToString();
			if (Funcion.iEscalarSQL(cdsCuenta, stSelect, true) > 0)
			{
				stSelect = "Select idCompra From Compra Where idAsiento = " + idAsi.ToString();
				int idCompra = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select idTipoFactura From Compra Where idAsiento = " + idAsi.ToString();
				int IdTipoF = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
//				Compra miCompra = new Compra(IdTipoF, idCompra);
//				miCompra.MdiParent = MdiParent;
//				miCompra.Show();
				return;
			}
			// Busca en factura generada por lotes
			stSelect = "Select Count(*) From Compra Where idAsientoPer = " + idAsi.ToString();
			if (Funcion.iEscalarSQL(cdsCuenta, stSelect, true) > 0)
			{
				stSelect = "Select top 1 idCompra From Compra Where idAsientoPer = " + idAsi.ToString();
				int idCompra = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select idTipoFactura From Compra Where idCompra = " + idCompra.ToString();
				int IdTipoF = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				MessageBox.Show("Asiento de Factura por lotes generada", "Información");
//				Compra miCompra = new Compra(IdTipoF, idCompra);
//				miCompra.MdiParent = MdiParent;
//				miCompra.Show();
				return;
			}

			// Busca Retencion de factura
			stSelect = "Select Count(*) From Compra Where idAsiRet = " + idAsi.ToString();
			if (Funcion.iEscalarSQL(cdsCuenta, stSelect, true) > 0)
			{
				stSelect = "Select top 1 idCompra From Compra Where idAsiRet = " + idAsi.ToString();
				int idCompra = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select idTipoFactura From Compra Where idCompra = " + idCompra.ToString();
				int IdTipoF = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				MessageBox.Show("Asiento de Retencion", "Información");
//				Compra miCompra = new Compra(IdTipoF, idCompra);
//				miCompra.MdiParent = MdiParent;
//				miCompra.Show();
				return;
			}

			// Busca Pago por Lotes
			stSelect = "Select Count(*) From PagoLote Where idAsiento = " + idAsi.ToString();
			if (Funcion.iEscalarSQL(cdsCuenta, stSelect, true) > 0)
			{
				stSelect = "Select top 1 IsNull(idCompra, 0) From Pago Where idAsiento = " + idAsi.ToString();
				int idCompra = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select top 1 idPagoLote From PagoLote Where idAsiento = " + idAsi.ToString();
				int idPagoLote = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select top 1 IsNull(idDeposito, 0) From PagoLote Where idAsiento = " + idAsi.ToString();
				int idDeposito = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);

				stSelect = "Select idTipoFactura From Compra Where idCompra = " + idCompra.ToString();
				int IdTipoF = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				if (idDeposito > 0)
				{
					MessageBox.Show("Asiento de Deposito de Pago por lotes", "Información");
					PagoLoteDia miPago = new PagoLoteDia(IdTipoF, idPagoLote);
					miPago.MdiParent = MdiParent;
					miPago.Show();
				}
				else
				{
					MessageBox.Show("Asiento de Pago por lotes", "Información");
					PagoLotes miPago = new PagoLotes(IdTipoF, idPagoLote);
					miPago.MdiParent = MdiParent;
					miPago.Show();
				}
				return;
			}	
			
			// Busca Pago individual
			stSelect = "Select Count(*) From Pago Where idAsiento = " + idAsi.ToString();
			if (Funcion.iEscalarSQL(cdsCuenta, stSelect, true) > 0)
			{
				stSelect = "Select top 1 idCompra From Pago Where idAsiento = " + idAsi.ToString();
				int idCompra = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				stSelect = "Select idTipoFactura From Compra Where idCompra = " + idCompra.ToString();
				int IdTipoF = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				MessageBox.Show("Asiento de Pago", "Información");
//				Pagos miPago = new Pagos(IdTipoF, idCompra);
//				miPago.MdiParent = MdiParent;
//				miPago.Show();
				return;
			}
			MessageBox.Show("Asiento No Existe en Modulos", "Información");
		}

		private void label11_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;

			if (cmbSubProyecto.Value == System.DBNull.Value || cmbSubProyecto.Text.Trim().Length == 0)
			{
				MessageBox.Show("No puede pasar proyectos vacios a las lineas", "Información");
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea pasar los proyectos a las lineas?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			int IdSubProyecto = (int) cmbSubProyecto.Value;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
			{
				if (dr.Cells["idSubProyecto"].Value == DBNull.Value 
					|| (int) dr.Cells["idSubProyecto"].Value == 0)
					dr.Cells["idSubProyecto"].Value = IdSubProyecto;
			}
			grAsiento.Refresh();
		}

		private void label10_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;

			if (cmbProyecto.Value == System.DBNull.Value || cmbProyecto.Text.Trim().Length == 0)
			{
				MessageBox.Show("No puede pasar proyectos vacios a las lineas", "Información");
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea pasar los proyectos a las lineas?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			int IdProyecto = (int) cmbProyecto.Value;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
			{
				if (dr.Cells["idProyecto"].Value == DBNull.Value 
					|| (int) dr.Cells["idProyecto"].Value == 0)
						dr.Cells["idProyecto"].Value = IdProyecto;
			}
			grAsiento.Refresh();
		}

		#endregion Barra Extendida
		#region cds
		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCuenta.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirConta);		
			string stSelect = "Select IsNull(ListarCuentas, 1) From SeteoC";
			bool bCargaArticulo = Funcion.bEscalarSQL(cdsCuenta, stSelect, true);
			string stFiltro = "idCuenta > 0";
			if (!bCargaArticulo) stFiltro = "idCuenta < 2";
			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsCuenta2"];
			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cuenta"], stFiltro));
		}

		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsProyecto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsMoneda_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsMoneda.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsCuentaCorr_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCuentaCorr.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsAsiento_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			if (e.DataMember.ToString() == "_Asiento")
			{
				drAsi = e.Row;
				lblAnulado.Visible = (bool) drAsi["Borrar"];
			}
		}

		#endregion cds
		#region Grilla
		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["idCuentaCod"].Value = e.Row.Cells["idCuenta"].Value;
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			////////////////cambios para BOTROSA proyecto por centro de costo
			////el siguiente codigo es para filtrar cc por proy
			if (Funcion.bEjecutaSQL(cdsAsiento, "Exec SeteoGExiste 'FTRPBP'")) 				
			{
				if (e.Cell.IsInEditMode)
				{
					if (e.Cell.Column.ToString() == "CentroCostoBR")
					{
						if(controlW != e.Cell.Row.Cells["CentroCostoBR"].Value.ToString())
						{
							controlCCP();
							controlW = e.Cell.Row.Cells["CentroCostoBR"].Value.ToString();							
							NuevoCodigo = codCC(controlW);
							NuevoCodigo=NuevoCodigo.Substring(0,5);		
							string stSelect = "Select idProyecto From Proyecto Where CentroCosto like '"
								+ NuevoCodigo + "%'";
							int IdProy = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
							e.Cell.Row.Cells["idProyecto"].Value = IdProy;
						}
						else return;
					}
				}
				if (e.Cell.IsInEditMode)
				{
					if (e.Cell.Column.ToString() == "ProyectoBR")
					{
						if(grAsiento.Rows.Count < 1 )return;
						string CodCentroCosto = grAsiento.ActiveRow.Cells["CentroCostoBR"].Value.ToString();							
						string NuevoCodigo1 = codCC(CodCentroCosto);
						NuevoCodigo1=NuevoCodigo1.Substring(0,5);	

						if(controlW1 != e.Cell.Row.Cells["ProyectoBR"].Value.ToString())
						{
							controlCPP(NuevoCodigo1);							
							controlW1 = e.Cell.Row.Cells["ProyectoBR"].Value.ToString();

							string stSelect1 = "Select idSubProyecto From SubProyecto Where nombre like '"
								+ controlW1 + "%'";
							int IdProy1 = Funcion.iEscalarSQL(cdsAsiento, stSelect1, true);
							e.Cell.Row.Cells["idSubProyecto"].Value = IdProy1;
						}
						else return;
					}
				}
			}
			/////			
			////////////////


			if (!bActualiza) return;

			if (e.Cell.IsInEditMode)
				if (e.Cell.Column.ToString() == "idCuentaCod")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["idCuenta"].Value = e.Cell.Row.Cells["idCuentaCod"].Value;
					}

			bActualiza = false;

			if (e.Cell.Column.ToString() == "Debe" || e.Cell.Column.ToString() == "Haber")
			{
				double dDebe = (double) grAsiento.ActiveRow.Cells["Debe"].Value;
				double dHaber = (double) grAsiento.ActiveRow.Cells["Haber"].Value;
				if (e.Cell.Column.ToString() == "Debe" && dHaber > 0 && dDebe > 0)
					grAsiento.ActiveRow.Cells["Haber"].Value = 0;
				if (e.Cell.Column.ToString() == "Haber" && dDebe > 0 && dHaber > 0)
					grAsiento.ActiveRow.Cells["Debe"].Value = 0;

				CalculoTotal();
			}
			#region Codigo o descripcion
			if (e.Cell.Column.ToString() == "Codigo")
			{
				string stCod = e.Cell.Row.Cells["Codigo"].Value.ToString();
				string stSelect = "Select Count(*) From Cuenta Where Codigo = '" + stCod + "'";
				int iCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
				int IdCuenta = 0;
				if (iCuenta != 1)
				{
					using (BuscaCuentas miBusqueda = new BuscaCuentas(stCod, 1))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							IdCuenta = (int) miBusqueda.txtIdBusca.Value;
							e.Cell.Row.Cells["Descripcion"].Value = miBusqueda.txtDescripcion.Text;
							e.Cell.Row.Cells["Codigo"].Value = miBusqueda.txtCodigo.Text;
						}
						else
						{
							bActualiza = true;
							return;						
						}
					}
				}
				else
				{
					stSelect = "Select idCuenta From Cuenta Where Codigo = '" + stCod + "'";
					IdCuenta = Funcion.iEscalarSQL(cdsCuenta, stSelect, true);
					stSelect = "Select Codigo From Cuenta Where Codigo = '" + stCod + "'";
					e.Cell.Row.Cells["Codigo"].Value = Funcion.sEscalarSQL(cdsCuenta, stSelect, true);
					stSelect = "Select Descripcion From Cuenta Where Codigo = '" + stCod + "'";
					e.Cell.Row.Cells["Descripcion"].Value = Funcion.sEscalarSQL(cdsCuenta, stSelect, true);
				}
				e.Cell.Row.Cells["idCuenta"].Value = IdCuenta;
				CalculoTotal();
			}
			#endregion Codigo o descripcion
			bActualiza = true;
		}

		private void grAsiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.Alt) 
				return;
			if (grAsiento.ActiveCell == null) return;
			switch(e.KeyValue)
			{
				case (int)Keys.F3:
					if ((this.grAsiento.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" || this.grAsiento.ActiveCell.Column.Key.ToString().ToUpper() == "DESCRIPCION") && barraDato1.bEditar)
					{
						using (BuscaCuentas miBusqueda = new BuscaCuentas("", 1))
						{
							if (DialogResult.OK == miBusqueda.ShowDialog())
							{								
								this.grAsiento.ActiveRow.Cells["Codigo"].Value = miBusqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value;// miBusqueda.txtCodigo.Text;
							}
							else
							{
								bActualiza = true;
								return;						
							}
						}
					}
					break;
				case (int) Keys.PageDown:
					if(grAsiento.ActiveCell.DroppedDown == false)
						grAsiento.ActiveCell.DroppedDown = true;
					break;

				case (int) Keys.Up:
					if (grAsiento.ActiveCell.DroppedDown) return;

					grAsiento.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grAsiento.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					grAsiento.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (grAsiento.ActiveCell.DroppedDown) return;
			
					grAsiento.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grAsiento.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					grAsiento.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.End:
					string stCol = grAsiento.ActiveCell.Column.ToString();
					if (!(stCol == "Debe" || stCol == "Haber")) return;

					double lnDiferencia = 0;
					if (stCol == "Haber")
						lnDiferencia = (double) txtTotalD.Value - (double) txtTotalH.Value;
					if (stCol== "Debe")
						lnDiferencia = -(double) txtTotalD.Value + (double) txtTotalH.Value;
					if (grAsiento.ActiveCell.Value != System.DBNull.Value)
						lnDiferencia += (double) grAsiento.ActiveCell.Value;
					lnDiferencia = Math.Round(lnDiferencia, 2);
					grAsiento.ActiveCell.Value = lnDiferencia;
					break;

				case (int) Keys.Enter:
					grAsiento.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grAsiento.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					grAsiento.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
			}		
		}

		private void grAsiento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			C1DataRow dr = cdsSeteoC.TableViews["SeteoC"].Rows[0];
			drSeteoC = SeteoCRow_tableView.Obj(dr);
			if ( drSeteoC.VerDeposito == false)
				e.Layout.Bands[0].Columns["Deposito"].Hidden = true;
			e.Layout.Bands[0].Columns["idProyecto"].Hidden = !(bool) drSeteoC.ProyectoLinea;
			e.Layout.Bands[0].Columns["idSubProyecto"].Hidden = !(bool) drSeteoC.SubProyLinea;
			if ( drSeteoC.verFechaComp == false)
				e.Layout.Bands[0].Columns["FechaComp"].Hidden = true;

			if (!drSeteoC.IsListarCuentasNull() && !(bool) drSeteoC.ListarCuentas)
			{
				e.Layout.Bands[0].Columns["Descripcion"].Hidden = false;
				e.Layout.Bands[0].Columns["Codigo"].Hidden = false;
				e.Layout.Bands[0].Columns["idCuentaCod"].Hidden = true;
				e.Layout.Bands[0].Columns["idCuenta"].Hidden = true;
			}
		}

		private void grAsiento_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.No ==	MessageBox.Show("¿Desea Borrar Lineas Seleccionadas?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				e.Cancel = true;
			else
				e.DisplayPromptMsg = false;
		}
		#endregion Grilla
		#region Botones Pantalla Principales
		private void CalculoTotal()
		{
			double dDebe = 0;
			double dHaber = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
			{
				if (dr.Cells["Debe"].Value != System.DBNull.Value)
					dDebe += (double) dr.Cells["Debe"].Value;
				if (dr.Cells["Haber"].Value != System.DBNull.Value)
					dHaber += (double) dr.Cells["Haber"].Value;
			}
			drAsi["Total"] = Math.Round(dDebe,2);
			txtTotalD.Value = Math.Round(dDebe,2);
			drAsi["TotalH"] = Math.Round(dHaber,2);
			txtTotalH.Value = Math.Round(dHaber,2);
		}

		private void cmbMoneda_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Moneda miTipoMoneda= new Moneda();
			miTipoMoneda.Show();
		}

		private void cmbProyecto_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Middle)
			{
				using (Proyecto miProyecto = new Proyecto())
				{
					miProyecto.ShowDialog();
				}
				cdsProyecto.Clear();
				cdsProyecto.Fill();
				return;
			}
			if (e.Button == MouseButtons.Right) 
			{
				if (grAsiento.ActiveRow != null)
					grAsiento.ActiveRow.Cells["idProyecto"].Value = 0;
				cmbProyecto.Value = 0;
				cmbProyecto.Refresh();
			}
		}

		private void cmbSubProyecto_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Middle)
			{
				using (SubProyecto miSubProyecto = new SubProyecto())
				{
					miSubProyecto.ShowDialog();
				}
				cdsProyecto.Clear();
				cdsProyecto.Fill();
				return;
			}
			if (e.Button == MouseButtons.Right)
			{
				if (grAsiento.ActiveRow != null)
					grAsiento.ActiveRow.Cells["idSubProyecto"].Value = 0;
				cmbSubProyecto.Value = 0;
				cmbSubProyecto.Refresh();
			}
		}

		private void cmbTipo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			AsientoTipo miTipoAsi= new AsientoTipo();
			miTipoAsi.Show();		
		}

		private void txtNumero_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (!(e.Button == MouseButtons.Right || e.Button == MouseButtons.Middle)) return;
			string stNumero = txtNumero.Text.Trim();
			if (stNumero.Length == 0) return;
			//Busqueda de compras
			Cursor = Cursors.WaitCursor;
			string stSelect = "Select Top 1 idCompra From Compra Where Numero = '"
				+ stNumero + "'";
			int IdCompra = Funcion.iEscalarSQL(cdsAsiento, stSelect, false);
			stSelect = "Select idTipoFactura From Compra Where idCompra = "
				+ IdCompra.ToString();
			int IdTipoFactura = Funcion.iEscalarSQL(cdsAsiento, stSelect, false);
			if (IdTipoFactura > 0 && IdCompra > 0)
			{
//				Compra miCompra = new Compra(IdTipoFactura, IdCompra);
//				miCompra.MdiParent = MdiParent;
//				miCompra.Show();
			}
			else
			{
				MessageBox.Show("Factura no encontrada", "Información");
			}
			Cursor = Cursors.Default;		
		}


		#endregion Botones Principales
		private int CuentaCorrIdCuenta()
		{
			bool bEsCuentaC = false;
			int idCuenta = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
			{
				if (dr.Cells["Haber"].Value != System.DBNull.Value && (double)dr.Cells["Haber"].Value > 0)
				{
					if (dr.Cells["idCuenta"].Value != System.DBNull.Value && (int)dr.Cells["idCuenta"].Value > 0)
					{
						idCuenta = (int)dr.Cells["idCuenta"].Value;
						for (int i=0; i<iCuentaCuentasC; i++)
						{
							if (ListaCuentasC[i] == idCuenta)
							{
								return idCuenta;
							}
						}
						if (bEsCuentaC == true) break;
					}
				}
			}
			return idCuenta;
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			CalculoTotal();
//			double lnDiferencia = (double) txtTotalD.Value - (double) txtTotalH.Value;
//			if (Math.Abs(lnDiferencia) > .0001)
//			{
//				bOk = false; 
//				errorProvider.SetError(grAsiento, "Suma de Columna Debe y de Haber tienen una diferencia de " + lnDiferencia.ToString("#,##0.00"));
//			}

			if (txtDescripcion.Text.Trim().Length == 0)
			{
				errorProvider.SetError(txtDescripcion, "Ingrese Descripción");
				txtDescripcion.Focus();
				bOk = false;
			}
			// Verificacion de Cheque
			if (txtCheque.Text.ToString().Trim().Length > 0 && (int) cmbTipo.Value == 3)
			{
				int idCuentaC = CuentaCorrIdCuenta();
				string stString = String.Format("Exec VerificaCheque '{0}', {1}, {2}", 
					txtCheque.Text.ToString().Trim(), idCuentaC, drAsi["idAsiento"].ToString());
				int iCuentaCh = Funcion.iEscalarSQL(cdsAsiento, stString, true);
				if (iCuentaCh > 0) 
				{
					errorProvider.SetError(txtCheque, "Cheque ya existe con esta cuenta ");
					bOk = false;
				}		
			}
			// Cuando esta habilitado cambiar el periodo no realiza verificacion de fecha
			if (miAcceso.CambiarPeriodo) return bOk;

			// Verificacion de Fecha
			DateTime dtFechaAsi;
			dtFechaAsi = DateTime.Today.AddYears(-1000);
			if (dtFecha.Value.ToString().Trim().Length < 3)
			{
				errorProvider.SetError(dtFecha, "Ingrese Fecha");
				dtFecha.Focus();
				bOk = false;
			}
			else
			{
				dtFechaAsi = (DateTime) dtFecha.Value;
			}

			if (drSeteoC.IsFechaFinNull() || drSeteoC.IsFechaIniNull() || (dtFechaAsi < DateTime.Today.AddYears(-100)))
			{
				errorProvider.SetError(dtFecha, "");
			}
			else // Existen las tres fechas
			{
				if (dtFechaAsi > drSeteoC.FechaFin || dtFechaAsi < drSeteoC.FechaIni)
				{
					bOk = false;
					errorProvider.SetError(dtFecha, "Fecha debe estar entre el " 
						+ drSeteoC.FechaIni.ToString("dd/MMM/yyyy")+ " y el " 
						+ drSeteoC.FechaFin.ToString("dd/MMM/yyyy") + ".\n"
						+ "Fechas definidas en Archivo - Propiedades - Contabilidad");
					dtFecha.Focus();
				}
			}
			// Fin de verificacion de fecha

			return bOk;
		}
		
		private bool VerificaCodigo()
		{
			string stCodigo = "CodAsiento = '" + txtCodigo.Value+ "'";
			errorProvider.SetError(txtCodigo, "");
			if (txtCodigo.Text.Trim().Length == 0)
			{
				errorProvider.SetError(txtCodigo,"Ingrese Código");
				txtCodigo.Focus();
				return false;
			}

			if (Contabilidad.ExisteCodigo(cdsAsiento, 
				txtCodigo.Value.ToString().Trim(), 
				int.Parse(drAsi["idAsiento"].ToString())))
			{
				DataRow[] dr = cdsAsiento.StorageDataSet.Tables["Asiento"].Select(stCodigo);
				if (dr.Length > 0)
				{
					DateTime dtFechaAntes = (DateTime) dr[0]["Fecha"];
					errorProvider.SetError(txtCodigo, 
						"Codigo ya existe en " + dr[0]["Descripcion"].ToString().Trim()
						+ " de fecha " + dtFechaAntes.ToString("dd/MMM/yyyy"));
				}
				else
				{
					errorProvider.SetError(txtCodigo, 
						"Codigo ya existe ");
				}
				txtCodigo.Focus();
				return false;
			}
			return true;
		}

		private double ValorCheque()
		{
			double fValorCheque = 0;
			int iTipo = (int) cmbTipo.Value;

			string strDetalle = "idAsiento = " + drAsi["idAsiento"].ToString();
			string stDebeHaber = "Debe";
			if (int.Parse(cmbTipo.Value.ToString()) == 3) stDebeHaber = "Haber";
			string stSelect = "Select Top 1 " + stDebeHaber + " from asientodetalle "
				+ "inner join cuentacorriente "
				+ "on asientodetalle.idcuenta = cuentacorriente.idcuenta "
				+ "Where asientodetalle." + strDetalle + " And " 
				+ stDebeHaber + " > 0 Order by idDetAsiento";

			if (drSeteo["StoreProcedure"] != null && (bool) drSeteo["StoreProcedure"] == true)
        stSelect = "Exec AsientoChequeValor '" + stDebeHaber + "', " + drAsi["idAsiento"].ToString();

			try
			{
					fValorCheque = Funcion.dEscalarSQL(cdsAsiento, stSelect, false);
			}
			catch
			{
				fValorCheque = 0;
			}

			return fValorCheque;
		}

		private string[] CuentaBanco()
		{
			string [] stCuenta = new string[2];
			string strDetalle = "idAsiento = " + drAsi["idAsiento"].ToString();
			string stDebeHaber = "Debe";
			if (int.Parse(cmbTipo.Value.ToString()) == 3) stDebeHaber = "Haber";
			string stSelect = "Select CuentaCorriente.NumCuenta from asientodetalle "
				+ "inner join cuentacorriente "
				+ "on asientodetalle.idcuenta = cuentacorriente.idcuenta "
				+ "Where asientodetalle. " + strDetalle
        + " And " + stDebeHaber + " > 0";
			stCuenta[0] = Funcion.sEscalarSQL(cdsAsiento, stSelect, false);

			stSelect = "Select Banco.Nombre from asientodetalle "
				+ "inner join cuentacorriente "
				+ "on asientodetalle.idcuenta = cuentacorriente.idcuenta "
				+ "inner join Banco "
				+ "on Banco.idBanco = cuentacorriente.idBanco "
				+ "Where asientodetalle. " + strDetalle
				+ " And " + stDebeHaber + " > 0";
			stCuenta[1] = Funcion.sEscalarSQL(cdsAsiento, stSelect, false);

			if (stCuenta[0] == null) stCuenta[0] = "";
			if (stCuenta[1] == null) stCuenta[1] = "";
			return stCuenta;
		}


		#region Cambios para botrosa Proyecto por centro de costo
		private void controlProyecto()
		{
			try
			{
				string va;
				va = cmbProyecto.Value.ToString();				
				if (va.Length <= 0)
				{					
					return;
				}
				else
				{			
					string stExec = "SELECT CentroCosto FROM PROYECTO WHERE IdProyecto = " + va;
					string cod = Funcion.sEscalarSQL(cdsAsiento, stExec, true);
					cod.Substring(0,5);			
					cdvSubProyecto.RowFilter="CodSubProyecto like '"+ cod+"%'";
				}			
			}
			catch(Exception ex){Console.Write(ex);}
		}

		private void cmbProyecto_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (Funcion.bEjecutaSQL(cdsAsiento, "Exec SeteoGExiste 'FTRPBP'")) 				
			{				
				controlProyecto();
			}
		}

		private void controlCCP()
		{
			int iCuenta =0;
			string txtE = grAsiento.ActiveRow.Cells["CentroCostoBR"].Value.ToString();
			if (txtE.Length > 0)
			{
				try
				{
					string stSelect = "EXEC BuscaRCentroC '" + txtE + "',1";
					iCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
					if (iCuenta == 0)
					{
						MessageBox.Show("No Existen Registros Coincidentes");
						grAsiento.ActiveRow.Cells["CentroCostoBR"].Value = "";												                  
						return;                  
					}
					if (iCuenta > 0)
					{						
						string stEmpresa = "Select Nombre from Proyecto where idProyecto = "+iCuenta;
						string CentroC = Funcion.sEscalarSQL(cdsAsiento, stEmpresa, true);							
						controlW = CentroC;							
						grAsiento.ActiveRow.Cells["CentroCostoBR"].Value = CentroC;					
					}            
					if (iCuenta < 0)
					{
						using (BuscaCentroCosto miBusca = new BuscaCentroCosto(txtE))
						{
							if (DialogResult.OK == miBusca.ShowDialog())
							{
								try
								{
									int id = miBusca.idBusca;
									string stEmpresa1 = "Select Nombre from Proyecto where idProyecto = "+id;
									string CentroC1 = Funcion.sEscalarSQL(cdsAsiento, stEmpresa1, true);									
									controlW = CentroC1;									
									grAsiento.ActiveRow.Cells["CentroCostoBR"].Value = CentroC1;										
								}
								catch { }
							}
						}
					}			
				}
				catch (Exception ex) {Console.Write(ex); }			
			}
		}
		public string codCC(string cod)
		{
			string stCodCC = "Select CentroCosto from Proyecto where nombre like '%"+cod+"%'";
			string IdGlobal = Funcion.sEscalarSQL(cdsAsiento, stCodCC, true);	
			return IdGlobal;
		}
		private void controlCPP(string cod1)
		{
			int iCuenta =0;
			string txtE = grAsiento.ActiveRow.Cells["ProyectoBR"].Value.ToString();
			if (txtE.Length > 0)
			{
				try
				{
					string stSelect = "EXEC BuscaRProyecto '" + txtE + "',1,"+cod1;
					iCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelect, true);
					if (iCuenta == 0)
					{
						MessageBox.Show("No Existen Registros Coincidentes");
						grAsiento.ActiveRow.Cells["ProyectoBR"].Value = "";												                  
						return;                  
					}
					if (iCuenta > 0)
					{						
						string stEmpresa = "Select Nombre from SubProyecto where idSubProyecto = "+iCuenta;
						string CentroC = Funcion.sEscalarSQL(cdsAsiento, stEmpresa, true);							
						controlW1 = CentroC;							
						grAsiento.ActiveRow.Cells["ProyectoBR"].Value = CentroC;					
					}            
					if (iCuenta < 0)
					{
					
						using (BuscaCentroCosto miBusca1 = new BuscaCentroCosto(txtE,cod1))
						{
							if (DialogResult.OK == miBusca1.ShowDialog())
							{
								try
								{
									int id = miBusca1.idBusca;
									string stEmpresa1 = "Select Nombre from SubProyecto where idSubProyecto = "+id;
									string CentroC1 = Funcion.sEscalarSQL(cdsAsiento, stEmpresa1, true);									
									controlW1 = CentroC1;									
									grAsiento.ActiveRow.Cells["ProyectoBR"].Value = CentroC1;										
								}
								catch { }
							}
						}
					}			
				}
				catch (Exception ex) {Console.Write(ex); }			
			}
		}
		#endregion Botrosa

		private void btDigitalizacion_Click(object sender, System.EventArgs e)
		{
			IdAsiento = (int) drAsi["idAsiento"];
			int tipo=(int)cmbTipo.Value;
			//int tipo=1;
			using (CIngresos miImagen = new CIngresos(IdAsiento,tipo))
			{
				miImagen.ShowDialog();
			}
		}

		private void btDigitalizacion_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			using (ReporteDig miImagen = new ReporteDig())
			{
				miImagen.ShowDialog();
			}
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cmbTipo_ValueChanged(object sender, System.EventArgs e)
		{
			if (cmbTipo.Value == DBNull.Value) return;
			if ((int)cmbTipo.Value == 2)
			{
				lblCheque.Text = "Deposito";
			}
			else if ((int)cmbTipo.Value == 3)
			{
				lblCheque.Text = "Cheque";
			}
			else
				lblCheque.Text = "Chq Dep";
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void cmbTipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.dtFecha.Enabled) this.dtFecha.Focus();
				else this.txtDescripcion.Focus();
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDescripcion.Focus();
		}

		private void cmbTipo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtDescripcion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtDescripcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtCodigo.Enabled) this.txtCodigo.Focus();
				else this.cmbProyecto.Focus();
			}
		}

		private void cmbProyecto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNotas.Focus();
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				this.grAsiento.DisplayLayout.Bands[0].AddNew();
				this.grAsiento.ActiveCell = this.grAsiento.Rows[0].Cells["Codigo"];
			}
		}

		private void txtCodigo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbProyecto.Focus();
		}

		private void cmbTipo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!this.barraDato1.bEditar) return;

			if (this.cmbTipo.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Tipo de Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}

			if ((int)this.cmbTipo.Value == 3)
			{
				MessageBox.Show("No puede crear Egreso en este Modulo, Solo Desde Pagos Autorizados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}			
		}

		private void grAsiento_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (this.barraDato1.bEditar)
			{
				e.Row.Cells["idProyecto"].Value = 19;
			}
		}
	}
}
