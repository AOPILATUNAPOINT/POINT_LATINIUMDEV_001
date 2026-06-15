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
	/// Descripción breve de SeteoF.
	/// </summary>
	public class SeteoF : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor5;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor7;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor8;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor9;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor10;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor11;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor12;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor13;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor14;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor15;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor16;
		private C1.Data.C1DataSet cdsNumero;
		private C1.Data.C1DataView cdvNumero;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCostoManual;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaFin;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIni;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkImpideSaldoNegativo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSaldoBodega;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDecTotal;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor3;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor4;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor5;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVencimiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnNumLineas;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl6;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optListaPrecio;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkModificaPrecioV;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkModificaImpuesto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor6;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor18;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCodArticulo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor19;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVerificaRuc;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor20;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor21;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCopiaDist;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor22;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor23;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConsumoFact;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkIngrEgrInterno;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNotaVentaPrdef;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor7;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCotizaReserva;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDesplegable;
		private Infragistics.Win.Misc.UltraButton cmdListaD;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor24;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor26;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor27;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor28;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor29;
		private Infragistics.Win.Misc.UltraButton btNumero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnTiempoVenta;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor30;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkImportacion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor31;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor32;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGrupoReferencia;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor33;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor35;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsTipoAsiento;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoContado;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoCredito;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAutorizaPago;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor34;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor36;
		private Infragistics.Win.Misc.UltraButton btNoAsignar;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvSubProyecto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSubproyecto;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor8;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor37;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor38;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor39;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor40;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor41;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPagoInteres;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor42;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor43;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor25;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl7;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.GroupBox groupBox6;
		private System.Windows.Forms.GroupBox groupBox7;
		private System.Windows.Forms.GroupBox groupBox8;
		private System.Windows.Forms.GroupBox groupBox9;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkClienteRep;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkIceLinea;
		private Infragistics.Win.Misc.UltraButton btidClienteToNombre;
		private DevExpress.XtraEditors.LookUpEdit cmbCliente;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor17;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor44;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor45;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor46;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor47;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCombos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor49;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor50;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor51;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor52;
		private System.Windows.Forms.GroupBox groupBox10;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor48;
		private System.Windows.Forms.GroupBox groupBox11;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor53;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor54;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor55;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor56;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor57;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor58;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCredito;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDirEtiquetas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor59;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor9;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentajeGarantia;
		private Infragistics.Win.Misc.UltraButton btnIVA;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor60;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEncuesta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuponWeb;
		private System.ComponentModel.IContainer components;

		public SeteoF()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CajaChica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoAsiento");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraNumero", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abrev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Name");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transforma");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Copias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoContado", -1, "cmbTipoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsientoCredito", -1, "cmbTipoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto", -1, "cmbProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subproyecto", -1, "cmbSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargaSeriales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab6 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab7 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.groupBox7 = new System.Windows.Forms.GroupBox();
			this.chkPagoInteres = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.chkAutorizaPago = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor35 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCostoManual = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor25 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor26 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkGrupoReferencia = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.ultraCheckEditor5 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkImpideSaldoNegativo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSaldoBodega = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor40 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.btidClienteToNombre = new Infragistics.Win.Misc.UltraButton();
			this.ultraCheckEditor30 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor41 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor38 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.chkIceLinea = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmdListaD = new Infragistics.Win.Misc.UltraButton();
			this.chkDesplegable = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor20 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkVencimiento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor11 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor10 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor9 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor8 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor7 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkImportacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor21 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCodArticulo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor12 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor13 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor14 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor15 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor16 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbSubproyecto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvSubProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.btNoAsignar = new Infragistics.Win.Misc.UltraButton();
			this.cmbTipoCredito = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.udsTipoAsiento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipoContado = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.spnTiempoVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btNumero = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdvNumero = new C1.Data.C1DataView();
			this.cdsNumero = new C1.Data.C1DataSet();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.groupBox9 = new System.Windows.Forms.GroupBox();
			this.ultraCheckEditor44 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor17 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkClienteRep = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor43 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor32 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor31 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor22 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor36 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor34 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor33 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkVerificaRuc = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox8 = new System.Windows.Forms.GroupBox();
			this.btnIVA = new Infragistics.Win.Misc.UltraButton();
			this.label18 = new System.Windows.Forms.Label();
			this.ultraNumericEditor9 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbCliente = new DevExpress.XtraEditors.LookUpEdit();
			this.ultraCheckEditor23 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkNotaVentaPrdef = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.spnNumLineas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.ultraNumericEditor7 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFechaIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkCopiaDist = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label15 = new System.Windows.Forms.Label();
			this.ultraNumericEditor8 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.ultraNumericEditor5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.ultraNumericEditor4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.ultraNumericEditor3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.spnDecTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraTabPageControl6 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraCheckEditor60 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.txtPorcentajeGarantia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.txtCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.txtTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCheckEditor24 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.ultraNumericEditor6 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCheckEditor19 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraCheckEditor28 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor18 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor6 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.chkModificaImpuesto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkModificaPrecioV = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor4 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.optListaPrecio = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraTabPageControl7 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.chkEncuesta = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor59 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDirEtiquetas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox11 = new System.Windows.Forms.GroupBox();
			this.ultraCheckEditor57 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor56 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor53 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor54 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor55 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox10 = new System.Windows.Forms.GroupBox();
			this.ultraCheckEditor51 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCombos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor49 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor52 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor50 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor47 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor46 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor45 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor48 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor58 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox6 = new System.Windows.Forms.GroupBox();
			this.chkCotizaReserva = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor27 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor39 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor29 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkConsumoFact = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkIngrEgrInterno = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor37 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor42 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.chkCuponWeb = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl1.SuspendLayout();
			this.groupBox7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.groupBox5.SuspendLayout();
			this.groupBox4.SuspendLayout();
			this.ultraTabPageControl2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubproyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipoAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoContado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTiempoVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsNumero)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			this.groupBox9.SuspendLayout();
			this.groupBox8.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumLineas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDecTotal)).BeginInit();
			this.ultraTabPageControl6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentajeGarantia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optListaPrecio)).BeginInit();
			this.ultraTabPageControl7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDirEtiquetas)).BeginInit();
			this.groupBox11.SuspendLayout();
			this.groupBox10.SuspendLayout();
			this.groupBox6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.groupBox7);
			this.ultraTabPageControl1.Controls.Add(this.groupBox5);
			this.ultraTabPageControl1.Controls.Add(this.groupBox4);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(678, 339);
			// 
			// groupBox7
			// 
			this.groupBox7.Controls.Add(this.chkPagoInteres);
			this.groupBox7.Controls.Add(this.chkAutorizaPago);
			this.groupBox7.Controls.Add(this.ultraCheckEditor35);
			this.groupBox7.Location = new System.Drawing.Point(359, 186);
			this.groupBox7.Name = "groupBox7";
			this.groupBox7.Size = new System.Drawing.Size(272, 121);
			this.groupBox7.TabIndex = 35;
			this.groupBox7.TabStop = false;
			this.groupBox7.Text = "Pagos";
			// 
			// chkPagoInteres
			// 
			this.chkPagoInteres.BackColor = System.Drawing.Color.Transparent;
			this.chkPagoInteres.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkPagoInteres.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.PagoInteres"));
			this.chkPagoInteres.Location = new System.Drawing.Point(24, 78);
			this.chkPagoInteres.Name = "chkPagoInteres";
			this.chkPagoInteres.Size = new System.Drawing.Size(168, 21);
			this.chkPagoInteres.TabIndex = 33;
			this.chkPagoInteres.Text = "Pagos con interés";
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// chkAutorizaPago
			// 
			this.chkAutorizaPago.BackColor = System.Drawing.Color.Transparent;
			this.chkAutorizaPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkAutorizaPago.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.AutorizaPago"));
			this.chkAutorizaPago.Location = new System.Drawing.Point(24, 26);
			this.chkAutorizaPago.Name = "chkAutorizaPago";
			this.chkAutorizaPago.Size = new System.Drawing.Size(176, 21);
			this.chkAutorizaPago.TabIndex = 30;
			this.chkAutorizaPago.Text = "Pagos requieren Autorización";
			// 
			// ultraCheckEditor35
			// 
			this.ultraCheckEditor35.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor35.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor35.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.PagoNegativo"));
			this.ultraCheckEditor35.Location = new System.Drawing.Point(24, 52);
			this.ultraCheckEditor35.Name = "ultraCheckEditor35";
			this.ultraCheckEditor35.Size = new System.Drawing.Size(160, 21);
			this.ultraCheckEditor35.TabIndex = 29;
			this.ultraCheckEditor35.Text = "Permitir pago en Exceso";
			// 
			// groupBox5
			// 
			this.groupBox5.Controls.Add(this.ultraCheckEditor1);
			this.groupBox5.Controls.Add(this.ultraCheckEditor2);
			this.groupBox5.Controls.Add(this.ultraCheckEditor3);
			this.groupBox5.Controls.Add(this.chkCostoManual);
			this.groupBox5.Controls.Add(this.ultraCheckEditor25);
			this.groupBox5.Controls.Add(this.ultraCheckEditor26);
			this.groupBox5.Controls.Add(this.chkGrupoReferencia);
			this.groupBox5.Location = new System.Drawing.Point(47, 31);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(280, 233);
			this.groupBox5.TabIndex = 34;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "Asientos";
			// 
			// ultraCheckEditor1
			// 
			this.ultraCheckEditor1.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.AsientoAutomatico"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(24, 181);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(200, 21);
			this.ultraCheckEditor1.TabIndex = 8;
			this.ultraCheckEditor1.Text = "Asiento automático al Grabar";
			this.ultraCheckEditor1.Visible = false;
			// 
			// ultraCheckEditor2
			// 
			this.ultraCheckEditor2.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CopiarNota"));
			this.ultraCheckEditor2.Location = new System.Drawing.Point(24, 26);
			this.ultraCheckEditor2.Name = "ultraCheckEditor2";
			this.ultraCheckEditor2.Size = new System.Drawing.Size(232, 21);
			this.ultraCheckEditor2.TabIndex = 9;
			this.ultraCheckEditor2.Text = "Copia nota de factura a contabilidad";
			// 
			// ultraCheckEditor3
			// 
			this.ultraCheckEditor3.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor3.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor3.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CopiarNumero"));
			this.ultraCheckEditor3.Location = new System.Drawing.Point(24, 52);
			this.ultraCheckEditor3.Name = "ultraCheckEditor3";
			this.ultraCheckEditor3.Size = new System.Drawing.Size(232, 21);
			this.ultraCheckEditor3.TabIndex = 10;
			this.ultraCheckEditor3.Text = "Copia número de factura a contabilidad";
			// 
			// chkCostoManual
			// 
			this.chkCostoManual.BackColor = System.Drawing.Color.Transparent;
			this.chkCostoManual.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkCostoManual.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CostoManual"));
			this.chkCostoManual.Location = new System.Drawing.Point(24, 78);
			this.chkCostoManual.Name = "chkCostoManual";
			this.chkCostoManual.Size = new System.Drawing.Size(232, 21);
			this.chkCostoManual.TabIndex = 14;
			this.chkCostoManual.Text = "Asiento de costo manual al generar venta";
			this.chkCostoManual.CheckedChanged += new System.EventHandler(this.ultraCheckEditor17_CheckedChanged);
			// 
			// ultraCheckEditor25
			// 
			this.ultraCheckEditor25.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor25.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor25.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.AsientoFP"));
			this.ultraCheckEditor25.Location = new System.Drawing.Point(24, 103);
			this.ultraCheckEditor25.Name = "ultraCheckEditor25";
			this.ultraCheckEditor25.Size = new System.Drawing.Size(232, 22);
			this.ultraCheckEditor25.TabIndex = 32;
			this.ultraCheckEditor25.Text = "Asiento por Forma de Pago en Venta";
			// 
			// ultraCheckEditor26
			// 
			this.ultraCheckEditor26.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor26.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor26.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.AsientoFPC"));
			this.ultraCheckEditor26.Location = new System.Drawing.Point(24, 129);
			this.ultraCheckEditor26.Name = "ultraCheckEditor26";
			this.ultraCheckEditor26.Size = new System.Drawing.Size(232, 22);
			this.ultraCheckEditor26.TabIndex = 26;
			this.ultraCheckEditor26.Text = "Asiento por Forma de Pago en Compra";
			// 
			// chkGrupoReferencia
			// 
			this.chkGrupoReferencia.BackColor = System.Drawing.Color.Transparent;
			this.chkGrupoReferencia.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkGrupoReferencia.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.AgrupaReferencia"));
			this.chkGrupoReferencia.Location = new System.Drawing.Point(24, 155);
			this.chkGrupoReferencia.Name = "chkGrupoReferencia";
			this.chkGrupoReferencia.Size = new System.Drawing.Size(240, 22);
			this.chkGrupoReferencia.TabIndex = 28;
			this.chkGrupoReferencia.Text = "Asiento de Pago Agrupado por Referencia";
			// 
			// groupBox4
			// 
			this.groupBox4.Controls.Add(this.ultraCheckEditor5);
			this.groupBox4.Controls.Add(this.chkImpideSaldoNegativo);
			this.groupBox4.Controls.Add(this.chkSaldoBodega);
			this.groupBox4.Controls.Add(this.ultraCheckEditor40);
			this.groupBox4.Location = new System.Drawing.Point(359, 31);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(272, 138);
			this.groupBox4.TabIndex = 33;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "Inventario Negativo";
			// 
			// ultraCheckEditor5
			// 
			this.ultraCheckEditor5.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor5.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor5.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.AlertaSaldo"));
			this.ultraCheckEditor5.Location = new System.Drawing.Point(24, 26);
			this.ultraCheckEditor5.Name = "ultraCheckEditor5";
			this.ultraCheckEditor5.Size = new System.Drawing.Size(232, 21);
			this.ultraCheckEditor5.TabIndex = 12;
			this.ultraCheckEditor5.Text = "Alerta de saldo mínimo";
			// 
			// chkImpideSaldoNegativo
			// 
			this.chkImpideSaldoNegativo.BackColor = System.Drawing.Color.Transparent;
			this.chkImpideSaldoNegativo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkImpideSaldoNegativo.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.SaldoNegativo"));
			this.chkImpideSaldoNegativo.Location = new System.Drawing.Point(24, 52);
			this.chkImpideSaldoNegativo.Name = "chkImpideSaldoNegativo";
			this.chkImpideSaldoNegativo.Size = new System.Drawing.Size(232, 21);
			this.chkImpideSaldoNegativo.TabIndex = 13;
			this.chkImpideSaldoNegativo.Text = "Impide saldos negativos";
			this.chkImpideSaldoNegativo.CheckedChanged += new System.EventHandler(this.chkImpideSaldoNegativo_CheckedChanged);
			// 
			// chkSaldoBodega
			// 
			this.chkSaldoBodega.BackColor = System.Drawing.Color.Transparent;
			this.chkSaldoBodega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkSaldoBodega.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.SaldoPorBodega"));
			this.chkSaldoBodega.Location = new System.Drawing.Point(40, 78);
			this.chkSaldoBodega.Name = "chkSaldoBodega";
			this.chkSaldoBodega.Size = new System.Drawing.Size(216, 21);
			this.chkSaldoBodega.TabIndex = 17;
			this.chkSaldoBodega.Text = "Impide Saldo Negativo en Bodega";
			// 
			// ultraCheckEditor40
			// 
			this.ultraCheckEditor40.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor40.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor40.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.SaldoEnCotiza"));
			this.ultraCheckEditor40.Location = new System.Drawing.Point(24, 103);
			this.ultraCheckEditor40.Name = "ultraCheckEditor40";
			this.ultraCheckEditor40.Size = new System.Drawing.Size(224, 22);
			this.ultraCheckEditor40.TabIndex = 32;
			this.ultraCheckEditor40.Text = "Revisión de Saldos en Cotización";
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.groupBox3);
			this.ultraTabPageControl2.Controls.Add(this.groupBox2);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(678, 339);
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.btidClienteToNombre);
			this.groupBox3.Controls.Add(this.ultraCheckEditor30);
			this.groupBox3.Controls.Add(this.ultraCheckEditor41);
			this.groupBox3.Controls.Add(this.ultraCheckEditor38);
			this.groupBox3.Location = new System.Drawing.Point(424, 17);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(216, 121);
			this.groupBox3.TabIndex = 24;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Principal";
			// 
			// btidClienteToNombre
			// 
			this.btidClienteToNombre.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btidClienteToNombre.Location = new System.Drawing.Point(184, 24);
			this.btidClienteToNombre.Name = "btidClienteToNombre";
			this.btidClienteToNombre.Size = new System.Drawing.Size(24, 25);
			this.btidClienteToNombre.TabIndex = 22;
			this.btidClienteToNombre.Text = "...";
			this.btidClienteToNombre.Click += new System.EventHandler(this.btidClienteToNombre_Click);
			// 
			// ultraCheckEditor30
			// 
			this.ultraCheckEditor30.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor30.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor30.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerNombreLista"));
			this.ultraCheckEditor30.Location = new System.Drawing.Point(16, 26);
			this.ultraCheckEditor30.Name = "ultraCheckEditor30";
			this.ultraCheckEditor30.Size = new System.Drawing.Size(168, 21);
			this.ultraCheckEditor30.TabIndex = 17;
			this.ultraCheckEditor30.Text = "Lista Clientes y Proveedores";
			// 
			// ultraCheckEditor41
			// 
			this.ultraCheckEditor41.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor41.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor41.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerBodegaNombre"));
			this.ultraCheckEditor41.Location = new System.Drawing.Point(16, 52);
			this.ultraCheckEditor41.Name = "ultraCheckEditor41";
			this.ultraCheckEditor41.Size = new System.Drawing.Size(144, 21);
			this.ultraCheckEditor41.TabIndex = 21;
			this.ultraCheckEditor41.Text = "Nombre de Bodega";
			// 
			// ultraCheckEditor38
			// 
			this.ultraCheckEditor38.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor38.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor38.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerServicio"));
			this.ultraCheckEditor38.Location = new System.Drawing.Point(16, 78);
			this.ultraCheckEditor38.Name = "ultraCheckEditor38";
			this.ultraCheckEditor38.Size = new System.Drawing.Size(144, 21);
			this.ultraCheckEditor38.TabIndex = 20;
			this.ultraCheckEditor38.Text = "Servicio y Propina";
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.chkIceLinea);
			this.groupBox2.Controls.Add(this.cmdListaD);
			this.groupBox2.Controls.Add(this.chkDesplegable);
			this.groupBox2.Controls.Add(this.ultraCheckEditor20);
			this.groupBox2.Controls.Add(this.chkVencimiento);
			this.groupBox2.Controls.Add(this.ultraCheckEditor11);
			this.groupBox2.Controls.Add(this.ultraCheckEditor10);
			this.groupBox2.Controls.Add(this.ultraCheckEditor9);
			this.groupBox2.Controls.Add(this.ultraCheckEditor8);
			this.groupBox2.Controls.Add(this.ultraCheckEditor7);
			this.groupBox2.Controls.Add(this.chkImportacion);
			this.groupBox2.Controls.Add(this.ultraCheckEditor21);
			this.groupBox2.Controls.Add(this.chkCodArticulo);
			this.groupBox2.Controls.Add(this.ultraCheckEditor12);
			this.groupBox2.Controls.Add(this.ultraCheckEditor13);
			this.groupBox2.Controls.Add(this.ultraCheckEditor14);
			this.groupBox2.Controls.Add(this.ultraCheckEditor15);
			this.groupBox2.Controls.Add(this.ultraCheckEditor16);
			this.groupBox2.Location = new System.Drawing.Point(16, 17);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(392, 271);
			this.groupBox2.TabIndex = 23;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Grilla";
			// 
			// chkIceLinea
			// 
			this.chkIceLinea.BackColor = System.Drawing.Color.Transparent;
			this.chkIceLinea.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkIceLinea.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.IceLinea"));
			this.chkIceLinea.Location = new System.Drawing.Point(16, 232);
			this.chkIceLinea.Name = "chkIceLinea";
			this.chkIceLinea.Size = new System.Drawing.Size(144, 17);
			this.chkIceLinea.TabIndex = 19;
			this.chkIceLinea.Text = "Ice por línea";
			// 
			// cmdListaD
			// 
			this.cmdListaD.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmdListaD.Location = new System.Drawing.Point(144, 26);
			this.cmdListaD.Name = "cmdListaD";
			this.cmdListaD.Size = new System.Drawing.Size(24, 25);
			this.cmdListaD.TabIndex = 15;
			this.cmdListaD.Text = "...";
			this.cmdListaD.Click += new System.EventHandler(this.cmdListaD_Click);
			// 
			// chkDesplegable
			// 
			this.chkDesplegable.BackColor = System.Drawing.Color.Transparent;
			this.chkDesplegable.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkDesplegable.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerDesplegable"));
			this.chkDesplegable.Location = new System.Drawing.Point(16, 26);
			this.chkDesplegable.Name = "chkDesplegable";
			this.chkDesplegable.Size = new System.Drawing.Size(128, 21);
			this.chkDesplegable.TabIndex = 14;
			this.chkDesplegable.Text = "Listas de Artículos";
			this.toolTip1.SetToolTip(this.chkDesplegable, "Lista de Articulos ");
			this.chkDesplegable.CheckedChanged += new System.EventHandler(this.chkDesplegable_CheckedChanged);
			// 
			// ultraCheckEditor20
			// 
			this.ultraCheckEditor20.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor20.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor20.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.EscogerCodigo"));
			this.ultraCheckEditor20.Location = new System.Drawing.Point(16, 52);
			this.ultraCheckEditor20.Name = "ultraCheckEditor20";
			this.ultraCheckEditor20.Size = new System.Drawing.Size(160, 21);
			this.ultraCheckEditor20.TabIndex = 12;
			this.ultraCheckEditor20.Text = "Código de Artículo";
			// 
			// chkVencimiento
			// 
			this.chkVencimiento.BackColor = System.Drawing.Color.Transparent;
			this.chkVencimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkVencimiento.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerVencimiento"));
			this.chkVencimiento.Location = new System.Drawing.Point(208, 181);
			this.chkVencimiento.Name = "chkVencimiento";
			this.chkVencimiento.Size = new System.Drawing.Size(160, 21);
			this.chkVencimiento.TabIndex = 10;
			this.chkVencimiento.Text = "Fecha Vencimiento";
			// 
			// ultraCheckEditor11
			// 
			this.ultraCheckEditor11.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor11.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor11.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerUnidad"));
			this.ultraCheckEditor11.Location = new System.Drawing.Point(16, 155);
			this.ultraCheckEditor11.Name = "ultraCheckEditor11";
			this.ultraCheckEditor11.Size = new System.Drawing.Size(128, 22);
			this.ultraCheckEditor11.TabIndex = 4;
			this.ultraCheckEditor11.Text = "Unidades";
			// 
			// ultraCheckEditor10
			// 
			this.ultraCheckEditor10.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor10.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor10.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerRefNumero"));
			this.ultraCheckEditor10.Location = new System.Drawing.Point(16, 207);
			this.ultraCheckEditor10.Name = "ultraCheckEditor10";
			this.ultraCheckEditor10.Size = new System.Drawing.Size(144, 17);
			this.ultraCheckEditor10.TabIndex = 3;
			this.ultraCheckEditor10.Text = "Número de Referencia";
			// 
			// ultraCheckEditor9
			// 
			this.ultraCheckEditor9.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor9.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor9.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerRefCodigo"));
			this.ultraCheckEditor9.Location = new System.Drawing.Point(16, 181);
			this.ultraCheckEditor9.Name = "ultraCheckEditor9";
			this.ultraCheckEditor9.Size = new System.Drawing.Size(144, 17);
			this.ultraCheckEditor9.TabIndex = 2;
			this.ultraCheckEditor9.Text = "Código de Referencia";
			// 
			// ultraCheckEditor8
			// 
			this.ultraCheckEditor8.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor8.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor8.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerPeso"));
			this.ultraCheckEditor8.Location = new System.Drawing.Point(208, 207);
			this.ultraCheckEditor8.Name = "ultraCheckEditor8";
			this.ultraCheckEditor8.Size = new System.Drawing.Size(120, 21);
			this.ultraCheckEditor8.TabIndex = 1;
			this.ultraCheckEditor8.Text = "Peso";
			// 
			// ultraCheckEditor7
			// 
			this.ultraCheckEditor7.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor7.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor7.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerBodega"));
			this.ultraCheckEditor7.Location = new System.Drawing.Point(16, 78);
			this.ultraCheckEditor7.Name = "ultraCheckEditor7";
			this.ultraCheckEditor7.Size = new System.Drawing.Size(120, 21);
			this.ultraCheckEditor7.TabIndex = 0;
			this.ultraCheckEditor7.Text = "Bodega";
			// 
			// chkImportacion
			// 
			this.chkImportacion.BackColor = System.Drawing.Color.Transparent;
			this.chkImportacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkImportacion.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.Importacion"));
			this.chkImportacion.Location = new System.Drawing.Point(208, 155);
			this.chkImportacion.Name = "chkImportacion";
			this.chkImportacion.Size = new System.Drawing.Size(88, 22);
			this.chkImportacion.TabIndex = 18;
			this.chkImportacion.Text = "Importación";
			// 
			// ultraCheckEditor21
			// 
			this.ultraCheckEditor21.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor21.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor21.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerSubProyecto"));
			this.ultraCheckEditor21.Location = new System.Drawing.Point(208, 52);
			this.ultraCheckEditor21.Name = "ultraCheckEditor21";
			this.ultraCheckEditor21.Size = new System.Drawing.Size(120, 21);
			this.ultraCheckEditor21.TabIndex = 13;
			this.ultraCheckEditor21.Text = "Proyecto por Línea";
			// 
			// chkCodArticulo
			// 
			this.chkCodArticulo.BackColor = System.Drawing.Color.Transparent;
			this.chkCodArticulo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkCodArticulo.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerProyecto"));
			this.chkCodArticulo.Location = new System.Drawing.Point(208, 26);
			this.chkCodArticulo.Name = "chkCodArticulo";
			this.chkCodArticulo.Size = new System.Drawing.Size(160, 21);
			this.chkCodArticulo.TabIndex = 11;
			this.chkCodArticulo.Text = "Centro de Costo por Línea";
			// 
			// ultraCheckEditor12
			// 
			this.ultraCheckEditor12.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor12.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor12.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerFlete"));
			this.ultraCheckEditor12.Location = new System.Drawing.Point(208, 103);
			this.ultraCheckEditor12.Name = "ultraCheckEditor12";
			this.ultraCheckEditor12.Size = new System.Drawing.Size(120, 22);
			this.ultraCheckEditor12.TabIndex = 9;
			this.ultraCheckEditor12.Text = "Flete";
			// 
			// ultraCheckEditor13
			// 
			this.ultraCheckEditor13.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor13.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor13.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerNotas"));
			this.ultraCheckEditor13.Location = new System.Drawing.Point(208, 78);
			this.ultraCheckEditor13.Name = "ultraCheckEditor13";
			this.ultraCheckEditor13.Size = new System.Drawing.Size(136, 21);
			this.ultraCheckEditor13.TabIndex = 8;
			this.ultraCheckEditor13.Text = "Notas";
			// 
			// ultraCheckEditor14
			// 
			this.ultraCheckEditor14.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor14.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor14.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerDescPorc"));
			this.ultraCheckEditor14.Location = new System.Drawing.Point(16, 129);
			this.ultraCheckEditor14.Name = "ultraCheckEditor14";
			this.ultraCheckEditor14.Size = new System.Drawing.Size(152, 22);
			this.ultraCheckEditor14.TabIndex = 7;
			this.ultraCheckEditor14.Text = "Descuento en Porcentaje";
			// 
			// ultraCheckEditor15
			// 
			this.ultraCheckEditor15.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor15.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor15.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerDescuento"));
			this.ultraCheckEditor15.Location = new System.Drawing.Point(16, 103);
			this.ultraCheckEditor15.Name = "ultraCheckEditor15";
			this.ultraCheckEditor15.Size = new System.Drawing.Size(160, 22);
			this.ultraCheckEditor15.TabIndex = 6;
			this.ultraCheckEditor15.Text = "Descuento en Valor";
			// 
			// ultraCheckEditor16
			// 
			this.ultraCheckEditor16.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor16.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor16.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerPedido"));
			this.ultraCheckEditor16.Location = new System.Drawing.Point(208, 129);
			this.ultraCheckEditor16.Name = "ultraCheckEditor16";
			this.ultraCheckEditor16.Size = new System.Drawing.Size(120, 22);
			this.ultraCheckEditor16.TabIndex = 5;
			this.ultraCheckEditor16.Text = "Pedidos";
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.cmbSubproyecto);
			this.ultraTabPageControl3.Controls.Add(this.cmbProyecto);
			this.ultraTabPageControl3.Controls.Add(this.btNoAsignar);
			this.ultraTabPageControl3.Controls.Add(this.cmbTipoCredito);
			this.ultraTabPageControl3.Controls.Add(this.cmbTipoContado);
			this.ultraTabPageControl3.Controls.Add(this.ultraLabel3);
			this.ultraTabPageControl3.Controls.Add(this.spnTiempoVenta);
			this.ultraTabPageControl3.Controls.Add(this.btNumero);
			this.ultraTabPageControl3.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(678, 339);
			// 
			// cmbSubproyecto
			// 
			this.cmbSubproyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubproyecto.DataSource = this.cdvSubProyecto;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbSubproyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSubproyecto.DisplayMember = "Nombre";
			this.cmbSubproyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubproyecto.Location = new System.Drawing.Point(251, 122);
			this.cmbSubproyecto.Name = "cmbSubproyecto";
			this.cmbSubproyecto.Size = new System.Drawing.Size(136, 69);
			this.cmbSubproyecto.TabIndex = 10;
			this.cmbSubproyecto.ValueMember = "idSubProyecto";
			this.cmbSubproyecto.Visible = false;
			// 
			// cdvSubProyecto
			// 
			this.cdvSubProyecto.BindingContextCtrl = this;
			this.cdvSubProyecto.DataSet = this.cdsProyecto;
			this.cdvSubProyecto.RowFilter = "Visible = true";
			this.cdvSubProyecto.Sort = "Nombre";
			this.cdvSubProyecto.TableName = "";
			this.cdvSubProyecto.TableViewName = "SubProyecto";
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
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataSource = this.cdvProyecto;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Width = 149;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 8;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 9;
			ultraGridColumn19.Header.VisiblePosition = 10;
			ultraGridColumn20.Header.VisiblePosition = 11;
			ultraGridColumn21.Header.VisiblePosition = 12;
			ultraGridBand2.Columns.AddRange(new object[] {
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
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(32, 78);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(176, 86);
			this.cmbProyecto.TabIndex = 9;
			this.cmbProyecto.ValueMember = "idProyecto";
			this.cmbProyecto.Visible = false;
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible = true";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// btNoAsignar
			// 
			this.btNoAsignar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btNoAsignar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNoAsignar.Location = new System.Drawing.Point(168, 309);
			this.btNoAsignar.Name = "btNoAsignar";
			this.btNoAsignar.Size = new System.Drawing.Size(75, 25);
			this.btNoAsignar.TabIndex = 8;
			this.btNoAsignar.Text = "&No Asignar";
			this.btNoAsignar.Click += new System.EventHandler(this.btNoAsignar_Click);
			// 
			// cmbTipoCredito
			// 
			this.cmbTipoCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoCredito.DataSource = this.udsTipoAsiento;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbTipoCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipoCredito.DisplayMember = "TipoAsiento";
			this.cmbTipoCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoCredito.Location = new System.Drawing.Point(360, 155);
			this.cmbTipoCredito.Name = "cmbTipoCredito";
			this.cmbTipoCredito.Size = new System.Drawing.Size(144, 78);
			this.cmbTipoCredito.TabIndex = 7;
			this.cmbTipoCredito.ValueMember = "idTipoAsiento";
			this.cmbTipoCredito.Visible = false;
			// 
			// udsTipoAsiento
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsTipoAsiento.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn1,
																															 ultraDataColumn2});
			// 
			// cmbTipoContado
			// 
			this.cmbTipoContado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoContado.DataSource = this.udsTipoAsiento;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbTipoContado.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoContado.DisplayMember = "TipoAsiento";
			this.cmbTipoContado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoContado.Location = new System.Drawing.Point(184, 146);
			this.cmbTipoContado.Name = "cmbTipoContado";
			this.cmbTipoContado.Size = new System.Drawing.Size(144, 78);
			this.cmbTipoContado.TabIndex = 6;
			this.cmbTipoContado.ValueMember = "idTipoAsiento";
			this.cmbTipoContado.Visible = false;
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.ultraLabel3.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel3.Location = new System.Drawing.Point(375, 309);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(200, 25);
			this.ultraLabel3.TabIndex = 5;
			this.ultraLabel3.Text = "Tiempo que Muestra Totales y Vuelto";
			// 
			// spnTiempoVenta
			// 
			this.spnTiempoVenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.spnTiempoVenta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnTiempoVenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.TiempoVenta"));
			this.spnTiempoVenta.FormatString = "##";
			this.spnTiempoVenta.Location = new System.Drawing.Point(583, 309);
			this.spnTiempoVenta.MaskInput = "nn";
			this.spnTiempoVenta.MaxValue = 10;
			this.spnTiempoVenta.MinValue = 0;
			this.spnTiempoVenta.Name = "spnTiempoVenta";
			this.spnTiempoVenta.PromptChar = ' ';
			this.spnTiempoVenta.Size = new System.Drawing.Size(64, 22);
			this.spnTiempoVenta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnTiempoVenta.TabIndex = 4;
			// 
			// btNumero
			// 
			this.btNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btNumero.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNumero.Location = new System.Drawing.Point(8, 309);
			this.btNumero.Name = "btNumero";
			this.btNumero.Size = new System.Drawing.Size(152, 25);
			this.btNumero.TabIndex = 1;
			this.btNumero.Text = "Numeración por estaciones";
			this.btNumero.Click += new System.EventHandler(this.btNumero_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvNumero;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn26.Header.VisiblePosition = 8;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Width = 164;
			ultraGridColumn28.Header.Caption = "Número";
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Width = 91;
			ultraGridColumn29.Header.Caption = "Código";
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Width = 92;
			ultraGridColumn30.Header.Caption = "Automático";
			ultraGridColumn30.Header.VisiblePosition = 3;
			ultraGridColumn30.Width = 118;
			ultraGridColumn31.Header.VisiblePosition = 4;
			ultraGridColumn31.Width = 98;
			ultraGridColumn32.Header.VisiblePosition = 9;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 6;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 7;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 10;
			ultraGridColumn36.Header.VisiblePosition = 11;
			ultraGridColumn37.Header.VisiblePosition = 5;
			ultraGridColumn38.Header.Caption = "Tipo Contado";
			ultraGridColumn38.Header.VisiblePosition = 12;
			ultraGridColumn38.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn39.Header.Caption = "Tipo Credito";
			ultraGridColumn39.Header.VisiblePosition = 13;
			ultraGridColumn39.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn40.Header.VisiblePosition = 14;
			ultraGridColumn40.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn41.Header.VisiblePosition = 15;
			ultraGridColumn41.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn42.Header.VisiblePosition = 16;
			ultraGridColumn43.Header.VisiblePosition = 17;
			ultraGridColumn44.Header.VisiblePosition = 18;
			ultraGridColumn45.Header.VisiblePosition = 19;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn45});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(664, 296);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// cdvNumero
			// 
			this.cdvNumero.BindingContextCtrl = this;
			this.cdvNumero.DataSet = this.cdsNumero;
			this.cdvNumero.TableName = "";
			this.cdvNumero.TableViewName = "CompraNumero";
			// 
			// cdsNumero
			// 
			this.cdsNumero.BindingContextCtrl = this;
			this.cdsNumero.DataLibrary = "LibFacturacion";
			this.cdsNumero.DataLibraryUrl = "";
			this.cdsNumero.DataSetDef = "dsCompraTabla";
			this.cdsNumero.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsNumero.Name = "cdsNumero";
			this.cdsNumero.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsNumero.SchemaDef = null;
			this.cdsNumero.BeforeFill += new C1.Data.FillEventHandler(this.cdsNumero_BeforeFill);
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.groupBox9);
			this.ultraTabPageControl4.Controls.Add(this.groupBox8);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(678, 339);
			// 
			// groupBox9
			// 
			this.groupBox9.Controls.Add(this.ultraCheckEditor44);
			this.groupBox9.Controls.Add(this.ultraCheckEditor17);
			this.groupBox9.Controls.Add(this.chkClienteRep);
			this.groupBox9.Controls.Add(this.ultraCheckEditor43);
			this.groupBox9.Controls.Add(this.ultraCheckEditor32);
			this.groupBox9.Controls.Add(this.ultraCheckEditor31);
			this.groupBox9.Controls.Add(this.ultraCheckEditor22);
			this.groupBox9.Controls.Add(this.ultraCheckEditor36);
			this.groupBox9.Controls.Add(this.ultraCheckEditor34);
			this.groupBox9.Controls.Add(this.ultraCheckEditor33);
			this.groupBox9.Controls.Add(this.chkVerificaRuc);
			this.groupBox9.Location = new System.Drawing.Point(371, 8);
			this.groupBox9.Name = "groupBox9";
			this.groupBox9.Size = new System.Drawing.Size(272, 328);
			this.groupBox9.TabIndex = 36;
			this.groupBox9.TabStop = false;
			this.groupBox9.Text = "Verificaciones";
			// 
			// ultraCheckEditor44
			// 
			this.ultraCheckEditor44.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor44.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor44.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.SerialNoExiste"));
			this.ultraCheckEditor44.Location = new System.Drawing.Point(24, 185);
			this.ultraCheckEditor44.Name = "ultraCheckEditor44";
			this.ultraCheckEditor44.Size = new System.Drawing.Size(224, 22);
			this.ultraCheckEditor44.TabIndex = 37;
			this.ultraCheckEditor44.Text = "Ingresa seriales inexistentes en ventas";
			// 
			// ultraCheckEditor17
			// 
			this.ultraCheckEditor17.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor17.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor17.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.FacturaManual"));
			this.ultraCheckEditor17.Location = new System.Drawing.Point(24, 212);
			this.ultraCheckEditor17.Name = "ultraCheckEditor17";
			this.ultraCheckEditor17.Size = new System.Drawing.Size(200, 22);
			this.ultraCheckEditor17.TabIndex = 36;
			this.ultraCheckEditor17.Text = "Facturación Manual";
			// 
			// chkClienteRep
			// 
			this.chkClienteRep.BackColor = System.Drawing.Color.Transparent;
			this.chkClienteRep.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkClienteRep.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.ClienteRepetido"));
			this.chkClienteRep.Location = new System.Drawing.Point(24, 26);
			this.chkClienteRep.Name = "chkClienteRep";
			this.chkClienteRep.Size = new System.Drawing.Size(232, 21);
			this.chkClienteRep.TabIndex = 35;
			this.chkClienteRep.Text = "Nombre de Cliente y Proveedor Repetido";
			// 
			// ultraCheckEditor43
			// 
			this.ultraCheckEditor43.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor43.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor43.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.LimiteTransforma"));
			this.ultraCheckEditor43.Location = new System.Drawing.Point(24, 265);
			this.ultraCheckEditor43.Name = "ultraCheckEditor43";
			this.ultraCheckEditor43.Size = new System.Drawing.Size(184, 21);
			this.ultraCheckEditor43.TabIndex = 34;
			this.ultraCheckEditor43.Text = "Transformación con cupo";
			// 
			// ultraCheckEditor32
			// 
			this.ultraCheckEditor32.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor32.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor32.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.ActualizaCosto"));
			this.ultraCheckEditor32.Location = new System.Drawing.Point(24, 104);
			this.ultraCheckEditor32.Name = "ultraCheckEditor32";
			this.ultraCheckEditor32.Size = new System.Drawing.Size(152, 22);
			this.ultraCheckEditor32.TabIndex = 27;
			this.ultraCheckEditor32.Text = "Actualiza Costo al Grabar";
			// 
			// ultraCheckEditor31
			// 
			this.ultraCheckEditor31.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor31.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor31.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.ActualizaSaldo"));
			this.ultraCheckEditor31.Location = new System.Drawing.Point(24, 78);
			this.ultraCheckEditor31.Name = "ultraCheckEditor31";
			this.ultraCheckEditor31.Size = new System.Drawing.Size(152, 21);
			this.ultraCheckEditor31.TabIndex = 26;
			this.ultraCheckEditor31.Text = "Actualiza Saldo al Grabar";
			// 
			// ultraCheckEditor22
			// 
			this.ultraCheckEditor22.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor22.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor22.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.ExigeVendedor"));
			this.ultraCheckEditor22.Location = new System.Drawing.Point(24, 291);
			this.ultraCheckEditor22.Name = "ultraCheckEditor22";
			this.ultraCheckEditor22.Size = new System.Drawing.Size(232, 22);
			this.ultraCheckEditor22.TabIndex = 20;
			this.ultraCheckEditor22.Text = "Exigir Vendedor al grabar Venta y Egreso";
			// 
			// ultraCheckEditor36
			// 
			this.ultraCheckEditor36.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor36.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor36.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.ExigeAutorizacion"));
			this.ultraCheckEditor36.Location = new System.Drawing.Point(24, 239);
			this.ultraCheckEditor36.Name = "ultraCheckEditor36";
			this.ultraCheckEditor36.Size = new System.Drawing.Size(224, 21);
			this.ultraCheckEditor36.TabIndex = 31;
			this.ultraCheckEditor36.Text = "Exigir Autorizaciones del SRI";
			// 
			// ultraCheckEditor34
			// 
			this.ultraCheckEditor34.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor34.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor34.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.SerieLinea"));
			this.ultraCheckEditor34.Location = new System.Drawing.Point(24, 158);
			this.ultraCheckEditor34.Name = "ultraCheckEditor34";
			this.ultraCheckEditor34.Size = new System.Drawing.Size(200, 22);
			this.ultraCheckEditor34.TabIndex = 30;
			this.ultraCheckEditor34.Text = "Seriales Horizontalmente en Ventas";
			// 
			// ultraCheckEditor33
			// 
			this.ultraCheckEditor33.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor33.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor33.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.OrdenC_ConsumoL"));
			this.ultraCheckEditor33.Location = new System.Drawing.Point(24, 131);
			this.ultraCheckEditor33.Name = "ultraCheckEditor33";
			this.ultraCheckEditor33.Size = new System.Drawing.Size(224, 22);
			this.ultraCheckEditor33.TabIndex = 29;
			this.ultraCheckEditor33.Text = "Orden de Compra - Consumo por línea";
			// 
			// chkVerificaRuc
			// 
			this.chkVerificaRuc.BackColor = System.Drawing.Color.Transparent;
			this.chkVerificaRuc.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkVerificaRuc.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VerificaRUC"));
			this.chkVerificaRuc.Location = new System.Drawing.Point(24, 52);
			this.chkVerificaRuc.Name = "chkVerificaRuc";
			this.chkVerificaRuc.Size = new System.Drawing.Size(160, 21);
			this.chkVerificaRuc.TabIndex = 18;
			this.chkVerificaRuc.Text = "Ingreso Obligatorio de RUC";
			// 
			// groupBox8
			// 
			this.groupBox8.Controls.Add(this.btnIVA);
			this.groupBox8.Controls.Add(this.label18);
			this.groupBox8.Controls.Add(this.ultraNumericEditor9);
			this.groupBox8.Controls.Add(this.cmbCliente);
			this.groupBox8.Controls.Add(this.ultraCheckEditor23);
			this.groupBox8.Controls.Add(this.chkNotaVentaPrdef);
			this.groupBox8.Controls.Add(this.label9);
			this.groupBox8.Controls.Add(this.spnNumLineas);
			this.groupBox8.Controls.Add(this.label14);
			this.groupBox8.Controls.Add(this.ultraNumericEditor7);
			this.groupBox8.Controls.Add(this.label13);
			this.groupBox8.Controls.Add(this.txtIva);
			this.groupBox8.Controls.Add(this.label10);
			this.groupBox8.Controls.Add(this.label2);
			this.groupBox8.Controls.Add(this.label1);
			this.groupBox8.Controls.Add(this.dtFechaIni);
			this.groupBox8.Controls.Add(this.dtFechaFin);
			this.groupBox8.Controls.Add(this.chkCopiaDist);
			this.groupBox8.Location = new System.Drawing.Point(35, 8);
			this.groupBox8.Name = "groupBox8";
			this.groupBox8.Size = new System.Drawing.Size(320, 328);
			this.groupBox8.TabIndex = 35;
			this.groupBox8.TabStop = false;
			this.groupBox8.Text = "Predefinidos";
			// 
			// btnIVA
			// 
			this.btnIVA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnIVA.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnIVA.Location = new System.Drawing.Point(144, 128);
			this.btnIVA.Name = "btnIVA";
			this.btnIVA.Size = new System.Drawing.Size(88, 25);
			this.btnIVA.TabIndex = 145;
			this.btnIVA.Text = "&I.V.A.";
			this.btnIVA.Click += new System.EventHandler(this.btnIVA_Click);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(16, 291);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(114, 16);
			this.label18.TabIndex = 144;
			this.label18.Text = "Máximo Transferencia";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraNumericEditor9
			// 
			this.ultraNumericEditor9.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.MaximoTransferencia"));
			this.ultraNumericEditor9.FormatString = "";
			this.ultraNumericEditor9.Location = new System.Drawing.Point(144, 288);
			this.ultraNumericEditor9.Name = "ultraNumericEditor9";
			this.ultraNumericEditor9.PromptChar = ' ';
			this.ultraNumericEditor9.Size = new System.Drawing.Size(48, 22);
			this.ultraNumericEditor9.TabIndex = 143;
			// 
			// cmbCliente
			// 
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSeteoF, "SeteoF.idCliente"));
			this.cmbCliente.Location = new System.Drawing.Point(112, 160);
			this.cmbCliente.Name = "cmbCliente";
			// 
			// cmbCliente.Properties
			// 
			this.cmbCliente.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbCliente.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 60, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.Ascending),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Código", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None),
																																																					 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("idCliente", "", 20, DevExpress.Utils.FormatType.None, "", false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None)});
			this.cmbCliente.Properties.DisplayMember = "Nombre";
			this.cmbCliente.Properties.MaxLength = 400;
			this.cmbCliente.Properties.NullText = "";
			this.cmbCliente.Properties.PopupWidth = 400;
			this.cmbCliente.Properties.ValueMember = "idCliente";
			this.cmbCliente.Size = new System.Drawing.Size(184, 20);
			this.cmbCliente.TabIndex = 142;
			this.cmbCliente.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCliente_MouseUp);
			// 
			// ultraCheckEditor23
			// 
			this.ultraCheckEditor23.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor23.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor23.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.VentaCredito"));
			this.ultraCheckEditor23.Location = new System.Drawing.Point(24, 198);
			this.ultraCheckEditor23.Name = "ultraCheckEditor23";
			this.ultraCheckEditor23.Size = new System.Drawing.Size(184, 22);
			this.ultraCheckEditor23.TabIndex = 21;
			this.ultraCheckEditor23.Text = "Ventas a Crédito";
			// 
			// chkNotaVentaPrdef
			// 
			this.chkNotaVentaPrdef.BackColor = System.Drawing.Color.Transparent;
			this.chkNotaVentaPrdef.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkNotaVentaPrdef.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.NotaVentaP"));
			this.chkNotaVentaPrdef.Location = new System.Drawing.Point(24, 224);
			this.chkNotaVentaPrdef.Name = "chkNotaVentaPrdef";
			this.chkNotaVentaPrdef.Size = new System.Drawing.Size(208, 22);
			this.chkNotaVentaPrdef.TabIndex = 21;
			this.chkNotaVentaPrdef.Text = "Nota de Venta Predefinida en Venta";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(24, 95);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(96, 16);
			this.label9.TabIndex = 12;
			this.label9.Text = "Líneas por Factura";
			// 
			// spnNumLineas
			// 
			this.spnNumLineas.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.spnNumLineas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.NumeroLineas"));
			this.spnNumLineas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnNumLineas.Location = new System.Drawing.Point(144, 95);
			this.spnNumLineas.MaskInput = "nnnn";
			this.spnNumLineas.MinValue = 1;
			this.spnNumLineas.Name = "spnNumLineas";
			this.spnNumLineas.PromptChar = ' ';
			this.spnNumLineas.Size = new System.Drawing.Size(88, 22);
			this.spnNumLineas.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnNumLineas.TabIndex = 11;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(24, 128);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(22, 16);
			this.label14.TabIndex = 25;
			this.label14.Text = "ICE";
			// 
			// ultraNumericEditor7
			// 
			this.ultraNumericEditor7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.Ice"));
			this.ultraNumericEditor7.FormatString = "";
			this.ultraNumericEditor7.Location = new System.Drawing.Point(64, 128);
			this.ultraNumericEditor7.MaskInput = "nn.nn";
			this.ultraNumericEditor7.Name = "ultraNumericEditor7";
			this.ultraNumericEditor7.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor7.PromptChar = ' ';
			this.ultraNumericEditor7.Size = new System.Drawing.Size(48, 22);
			this.ultraNumericEditor7.TabIndex = 24;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(240, 24);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(26, 16);
			this.label13.TabIndex = 23;
			this.label13.Text = "IVA:";
			this.label13.Visible = false;
			// 
			// txtIva
			// 
			this.txtIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.Iva"));
			this.txtIva.FormatString = "";
			this.txtIva.Location = new System.Drawing.Point(240, 48);
			this.txtIva.MaskInput = "nn.nn";
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(48, 22);
			this.txtIva.TabIndex = 22;
			this.txtIva.Visible = false;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(24, 164);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(77, 16);
			this.label10.TabIndex = 14;
			this.label10.Text = "Cliente Predef.";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(24, 26);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(67, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Fecha Inicial";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(24, 60);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(60, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Fecha Final";
			// 
			// dtFechaIni
			// 
			this.dtFechaIni.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaIni.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.FechaInicio"));
			dateButton1.Caption = "Today";
			this.dtFechaIni.DateButtons.Add(dateButton1);
			this.dtFechaIni.Format = "dd/MMM/yyyy";
			this.dtFechaIni.Location = new System.Drawing.Point(120, 26);
			this.dtFechaIni.Name = "dtFechaIni";
			this.dtFechaIni.NonAutoSizeHeight = 23;
			this.dtFechaIni.Size = new System.Drawing.Size(112, 21);
			this.dtFechaIni.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtFechaIni.SpinButtonsVisible = true;
			this.dtFechaIni.TabIndex = 1;
			this.dtFechaIni.Value = new System.DateTime(2006, 2, 4, 0, 0, 0, 0);
			// 
			// dtFechaFin
			// 
			this.dtFechaFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaFin.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.FechaFinal"));
			dateButton2.Caption = "Today";
			this.dtFechaFin.DateButtons.Add(dateButton2);
			this.dtFechaFin.Format = "dd/MMM/yyyy";
			this.dtFechaFin.Location = new System.Drawing.Point(120, 60);
			this.dtFechaFin.Name = "dtFechaFin";
			this.dtFechaFin.NonAutoSizeHeight = 23;
			this.dtFechaFin.Size = new System.Drawing.Size(112, 21);
			this.dtFechaFin.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtFechaFin.SpinButtonsVisible = true;
			this.dtFechaFin.TabIndex = 0;
			this.dtFechaFin.Value = new System.DateTime(2006, 2, 4, 0, 0, 0, 0);
			// 
			// chkCopiaDist
			// 
			this.chkCopiaDist.BackColor = System.Drawing.Color.Transparent;
			this.chkCopiaDist.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkCopiaDist.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.copia_dist"));
			this.chkCopiaDist.Location = new System.Drawing.Point(24, 250);
			this.chkCopiaDist.Name = "chkCopiaDist";
			this.chkCopiaDist.Size = new System.Drawing.Size(280, 21);
			this.chkCopiaDist.TabIndex = 19;
			this.chkCopiaDist.Text = "Copiar Vendedor y Descuento de datos del Cliente";
			this.toolTip1.SetToolTip(this.chkCopiaDist, "Vendedor, descuento, dias de credito");
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.label15);
			this.ultraTabPageControl5.Controls.Add(this.ultraNumericEditor8);
			this.ultraTabPageControl5.Controls.Add(this.label8);
			this.ultraTabPageControl5.Controls.Add(this.ultraNumericEditor5);
			this.ultraTabPageControl5.Controls.Add(this.label7);
			this.ultraTabPageControl5.Controls.Add(this.ultraNumericEditor4);
			this.ultraTabPageControl5.Controls.Add(this.label6);
			this.ultraTabPageControl5.Controls.Add(this.ultraNumericEditor3);
			this.ultraTabPageControl5.Controls.Add(this.label5);
			this.ultraTabPageControl5.Controls.Add(this.ultraNumericEditor2);
			this.ultraTabPageControl5.Controls.Add(this.label4);
			this.ultraTabPageControl5.Controls.Add(this.ultraNumericEditor1);
			this.ultraTabPageControl5.Controls.Add(this.label3);
			this.ultraTabPageControl5.Controls.Add(this.spnDecTotal);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(678, 339);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(72, 240);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(60, 16);
			this.label15.TabIndex = 13;
			this.label15.Text = "Cant. Prod.";
			// 
			// ultraNumericEditor8
			// 
			appearance6.TextHAlign = Infragistics.Win.HAlign.Left;
			this.ultraNumericEditor8.ButtonAppearance = appearance6;
			this.ultraNumericEditor8.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.DecProduccion"));
			this.ultraNumericEditor8.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor8.FormatString = "##";
			this.ultraNumericEditor8.Location = new System.Drawing.Point(152, 240);
			this.ultraNumericEditor8.MaskInput = "nn";
			this.ultraNumericEditor8.MaxValue = 10;
			this.ultraNumericEditor8.MinValue = 0;
			this.ultraNumericEditor8.Name = "ultraNumericEditor8";
			this.ultraNumericEditor8.PromptChar = ' ';
			this.ultraNumericEditor8.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ultraNumericEditor8.Size = new System.Drawing.Size(48, 22);
			this.ultraNumericEditor8.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor8.TabIndex = 12;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(72, 208);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(29, 16);
			this.label8.TabIndex = 11;
			this.label8.Text = "Total";
			// 
			// ultraNumericEditor5
			// 
			appearance7.TextHAlign = Infragistics.Win.HAlign.Left;
			this.ultraNumericEditor5.ButtonAppearance = appearance7;
			this.ultraNumericEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.DecTotal"));
			this.ultraNumericEditor5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor5.FormatString = "##";
			this.ultraNumericEditor5.Location = new System.Drawing.Point(152, 208);
			this.ultraNumericEditor5.MaskInput = "nn";
			this.ultraNumericEditor5.MaxValue = 10;
			this.ultraNumericEditor5.MinValue = 0;
			this.ultraNumericEditor5.Name = "ultraNumericEditor5";
			this.ultraNumericEditor5.PromptChar = ' ';
			this.ultraNumericEditor5.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ultraNumericEditor5.Size = new System.Drawing.Size(48, 22);
			this.ultraNumericEditor5.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor5.TabIndex = 10;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(72, 168);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(54, 16);
			this.label7.TabIndex = 9;
			this.label7.Text = "Redondeo";
			// 
			// ultraNumericEditor4
			// 
			appearance8.TextHAlign = Infragistics.Win.HAlign.Left;
			this.ultraNumericEditor4.ButtonAppearance = appearance8;
			this.ultraNumericEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.DecRedondeo"));
			this.ultraNumericEditor4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor4.FormatString = "##";
			this.ultraNumericEditor4.Location = new System.Drawing.Point(152, 168);
			this.ultraNumericEditor4.MaskInput = "nn";
			this.ultraNumericEditor4.MaxValue = 10;
			this.ultraNumericEditor4.MinValue = 0;
			this.ultraNumericEditor4.Name = "ultraNumericEditor4";
			this.ultraNumericEditor4.PromptChar = ' ';
			this.ultraNumericEditor4.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ultraNumericEditor4.Size = new System.Drawing.Size(48, 22);
			this.ultraNumericEditor4.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor4.TabIndex = 8;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(72, 136);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 7;
			this.label6.Text = "Precio";
			// 
			// ultraNumericEditor3
			// 
			appearance9.TextHAlign = Infragistics.Win.HAlign.Left;
			this.ultraNumericEditor3.ButtonAppearance = appearance9;
			this.ultraNumericEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.DecPrecio"));
			this.ultraNumericEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor3.FormatString = "##";
			this.ultraNumericEditor3.Location = new System.Drawing.Point(152, 136);
			this.ultraNumericEditor3.MaskInput = "nn";
			this.ultraNumericEditor3.MaxValue = 10;
			this.ultraNumericEditor3.MinValue = 0;
			this.ultraNumericEditor3.Name = "ultraNumericEditor3";
			this.ultraNumericEditor3.PromptChar = ' ';
			this.ultraNumericEditor3.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ultraNumericEditor3.Size = new System.Drawing.Size(48, 22);
			this.ultraNumericEditor3.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor3.TabIndex = 6;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(72, 104);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(71, 16);
			this.label5.TabIndex = 5;
			this.label5.Text = "% Descuento";
			// 
			// ultraNumericEditor2
			// 
			appearance10.TextHAlign = Infragistics.Win.HAlign.Left;
			this.ultraNumericEditor2.ButtonAppearance = appearance10;
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.DecPorcDesc"));
			this.ultraNumericEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor2.FormatString = "##";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(152, 104);
			this.ultraNumericEditor2.MaskInput = "nn";
			this.ultraNumericEditor2.MaxValue = 6;
			this.ultraNumericEditor2.MinValue = 0;
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ultraNumericEditor2.Size = new System.Drawing.Size(48, 22);
			this.ultraNumericEditor2.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor2.TabIndex = 4;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(72, 64);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(51, 16);
			this.label4.TabIndex = 3;
			this.label4.Text = "Impuesto";
			// 
			// ultraNumericEditor1
			// 
			appearance11.TextHAlign = Infragistics.Win.HAlign.Left;
			this.ultraNumericEditor1.ButtonAppearance = appearance11;
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.DecImpuesto"));
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.FormatString = "##";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(152, 64);
			this.ultraNumericEditor1.MaskInput = "nn";
			this.ultraNumericEditor1.MaxValue = 6;
			this.ultraNumericEditor1.MinValue = 0;
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ultraNumericEditor1.Size = new System.Drawing.Size(48, 22);
			this.ultraNumericEditor1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor1.TabIndex = 2;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(72, 32);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(48, 16);
			this.label3.TabIndex = 1;
			this.label3.Text = "Cantidad";
			// 
			// spnDecTotal
			// 
			appearance12.TextHAlign = Infragistics.Win.HAlign.Left;
			this.spnDecTotal.ButtonAppearance = appearance12;
			this.spnDecTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.DecCantidad"));
			this.spnDecTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnDecTotal.FormatString = "##";
			this.spnDecTotal.Location = new System.Drawing.Point(152, 32);
			this.spnDecTotal.MaskInput = "nn";
			this.spnDecTotal.MaxValue = 6;
			this.spnDecTotal.MinValue = 0;
			this.spnDecTotal.Name = "spnDecTotal";
			this.spnDecTotal.PromptChar = ' ';
			this.spnDecTotal.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.spnDecTotal.Size = new System.Drawing.Size(48, 22);
			this.spnDecTotal.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDecTotal.TabIndex = 0;
			// 
			// ultraTabPageControl6
			// 
			this.ultraTabPageControl6.Controls.Add(this.ultraCheckEditor60);
			this.ultraTabPageControl6.Controls.Add(this.label19);
			this.ultraTabPageControl6.Controls.Add(this.txtPorcentajeGarantia);
			this.ultraTabPageControl6.Controls.Add(this.label17);
			this.ultraTabPageControl6.Controls.Add(this.txtCredito);
			this.ultraTabPageControl6.Controls.Add(this.label16);
			this.ultraTabPageControl6.Controls.Add(this.txtTarjeta);
			this.ultraTabPageControl6.Controls.Add(this.ultraCheckEditor24);
			this.ultraTabPageControl6.Controls.Add(this.label12);
			this.ultraTabPageControl6.Controls.Add(this.ultraNumericEditor6);
			this.ultraTabPageControl6.Controls.Add(this.ultraCheckEditor19);
			this.ultraTabPageControl6.Controls.Add(this.groupBox1);
			this.ultraTabPageControl6.Controls.Add(this.chkModificaImpuesto);
			this.ultraTabPageControl6.Controls.Add(this.chkModificaPrecioV);
			this.ultraTabPageControl6.Controls.Add(this.ultraCheckEditor4);
			this.ultraTabPageControl6.Controls.Add(this.optListaPrecio);
			this.ultraTabPageControl6.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl6.Name = "ultraTabPageControl6";
			this.ultraTabPageControl6.Size = new System.Drawing.Size(678, 339);
			// 
			// ultraCheckEditor60
			// 
			this.ultraCheckEditor60.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor60.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor60.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.GarantiaExtendida"));
			this.ultraCheckEditor60.Location = new System.Drawing.Point(304, 168);
			this.ultraCheckEditor60.Name = "ultraCheckEditor60";
			this.ultraCheckEditor60.Size = new System.Drawing.Size(120, 22);
			this.ultraCheckEditor60.TabIndex = 40;
			this.ultraCheckEditor60.Text = "Garantia Extendida";
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(432, 168);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(119, 16);
			this.label19.TabIndex = 39;
			this.label19.Text = "Porcentaje de Garantia";
			// 
			// txtPorcentajeGarantia
			// 
			this.txtPorcentajeGarantia.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtPorcentajeGarantia.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.PorcentajeGarantiaExtendida"));
			this.txtPorcentajeGarantia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentajeGarantia.FormatString = "#,##0.00";
			this.txtPorcentajeGarantia.Location = new System.Drawing.Point(432, 192);
			this.txtPorcentajeGarantia.MaskInput = "nn,nnn,nnn";
			this.txtPorcentajeGarantia.MaxValue = 100;
			this.txtPorcentajeGarantia.MinValue = 0;
			this.txtPorcentajeGarantia.Name = "txtPorcentajeGarantia";
			this.txtPorcentajeGarantia.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentajeGarantia.PromptChar = ' ';
			this.txtPorcentajeGarantia.Size = new System.Drawing.Size(128, 22);
			this.txtPorcentajeGarantia.TabIndex = 38;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(304, 304);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(79, 16);
			this.label17.TabIndex = 32;
			this.label17.Text = "Tarjeta Crédito";
			this.label17.Visible = false;
			// 
			// txtCredito
			// 
			appearance13.TextHAlign = Infragistics.Win.HAlign.Left;
			this.txtCredito.ButtonAppearance = appearance13;
			this.txtCredito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.PorcContCredito"));
			this.txtCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCredito.FormatString = "#,##0.0000";
			this.txtCredito.Location = new System.Drawing.Point(424, 304);
			this.txtCredito.MaxValue = 100;
			this.txtCredito.MinValue = 0;
			this.txtCredito.Name = "txtCredito";
			this.txtCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCredito.PromptChar = ' ';
			this.txtCredito.Size = new System.Drawing.Size(136, 22);
			this.txtCredito.TabIndex = 31;
			this.txtCredito.Visible = false;
			this.txtCredito.Click += new System.EventHandler(this.txtCredito_Click);
			this.txtCredito.Enter += new System.EventHandler(this.txtCredito_Enter);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(304, 280);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(94, 16);
			this.label16.TabIndex = 30;
			this.label16.Text = "Contado a Tarjeta";
			this.label16.Visible = false;
			// 
			// txtTarjeta
			// 
			appearance14.TextHAlign = Infragistics.Win.HAlign.Left;
			this.txtTarjeta.ButtonAppearance = appearance14;
			this.txtTarjeta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.PorcContTarjeta"));
			this.txtTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTarjeta.FormatString = "#,##0.0000";
			this.txtTarjeta.Location = new System.Drawing.Point(424, 280);
			this.txtTarjeta.MaxValue = 100;
			this.txtTarjeta.MinValue = 0;
			this.txtTarjeta.Name = "txtTarjeta";
			this.txtTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTarjeta.PromptChar = ' ';
			this.txtTarjeta.Size = new System.Drawing.Size(136, 22);
			this.txtTarjeta.TabIndex = 29;
			this.txtTarjeta.Visible = false;
			this.txtTarjeta.Click += new System.EventHandler(this.txtTarjeta_Click);
			this.txtTarjeta.Enter += new System.EventHandler(this.txtTarjeta_Enter);
			// 
			// ultraCheckEditor24
			// 
			this.ultraCheckEditor24.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor24.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor24.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CambiarIvaC"));
			this.ultraCheckEditor24.Location = new System.Drawing.Point(24, 233);
			this.ultraCheckEditor24.Name = "ultraCheckEditor24";
			this.ultraCheckEditor24.Size = new System.Drawing.Size(232, 21);
			this.ultraCheckEditor24.TabIndex = 28;
			this.ultraCheckEditor24.Text = "Modificar Impuestos en Compra";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(304, 248);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(106, 16);
			this.label12.TabIndex = 27;
			this.label12.Text = "Límite de Descuento";
			this.label12.Visible = false;
			// 
			// ultraNumericEditor6
			// 
			appearance15.TextHAlign = Infragistics.Win.HAlign.Left;
			this.ultraNumericEditor6.ButtonAppearance = appearance15;
			this.ultraNumericEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.LimiteDescuento"));
			this.ultraNumericEditor6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor6.FormatString = "##";
			this.ultraNumericEditor6.Location = new System.Drawing.Point(440, 248);
			this.ultraNumericEditor6.MaskInput = "nnn";
			this.ultraNumericEditor6.MaxValue = 100;
			this.ultraNumericEditor6.MinValue = 0;
			this.ultraNumericEditor6.Name = "ultraNumericEditor6";
			this.ultraNumericEditor6.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor6.PromptChar = ' ';
			this.ultraNumericEditor6.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ultraNumericEditor6.Size = new System.Drawing.Size(64, 22);
			this.ultraNumericEditor6.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor6.TabIndex = 26;
			this.ultraNumericEditor6.Visible = false;
			// 
			// ultraCheckEditor19
			// 
			this.ultraCheckEditor19.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor19.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor19.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.PrecioTE"));
			this.ultraCheckEditor19.Location = new System.Drawing.Point(24, 258);
			this.ultraCheckEditor19.Name = "ultraCheckEditor19";
			this.ultraCheckEditor19.Size = new System.Drawing.Size(232, 22);
			this.ultraCheckEditor19.TabIndex = 25;
			this.ultraCheckEditor19.Text = "Precio / Costo en Egreso de Bodega";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.ultraCheckEditor28);
			this.groupBox1.Controls.Add(this.ultraCheckEditor18);
			this.groupBox1.Controls.Add(this.ultraCheckEditor6);
			this.groupBox1.Controls.Add(this.label11);
			this.groupBox1.Location = new System.Drawing.Point(304, 17);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(280, 147);
			this.groupBox1.TabIndex = 24;
			this.groupBox1.TabStop = false;
			// 
			// ultraCheckEditor28
			// 
			this.ultraCheckEditor28.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor28.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor28.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.c_u_c_p"));
			this.ultraCheckEditor28.Location = new System.Drawing.Point(32, 112);
			this.ultraCheckEditor28.Name = "ultraCheckEditor28";
			this.ultraCheckEditor28.Size = new System.Drawing.Size(200, 22);
			this.ultraCheckEditor28.TabIndex = 24;
			this.ultraCheckEditor28.Text = "Mostrar Costo último en Compra";
			// 
			// ultraCheckEditor18
			// 
			this.ultraCheckEditor18.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor18.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor18.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CompraUltimaC"));
			this.ultraCheckEditor18.Location = new System.Drawing.Point(32, 86);
			this.ultraCheckEditor18.Name = "ultraCheckEditor18";
			this.ultraCheckEditor18.Size = new System.Drawing.Size(224, 22);
			this.ultraCheckEditor18.TabIndex = 23;
			this.ultraCheckEditor18.Text = "Con Botón de Actualización de Precios";
			// 
			// ultraCheckEditor6
			// 
			this.ultraCheckEditor6.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor6.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor6.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.ActualizarCosto"));
			this.ultraCheckEditor6.Location = new System.Drawing.Point(32, 60);
			this.ultraCheckEditor6.Name = "ultraCheckEditor6";
			this.ultraCheckEditor6.Size = new System.Drawing.Size(144, 22);
			this.ultraCheckEditor6.TabIndex = 21;
			this.ultraCheckEditor6.Text = "Al Grabar la Compra";
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(16, 17);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(216, 35);
			this.label11.TabIndex = 22;
			this.label11.Text = "Actualiza Precio de Venta del Artículo en Base a Costo de Compra";
			// 
			// chkModificaImpuesto
			// 
			this.chkModificaImpuesto.BackColor = System.Drawing.Color.Transparent;
			this.chkModificaImpuesto.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkModificaImpuesto.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CambiarIva"));
			this.chkModificaImpuesto.Location = new System.Drawing.Point(24, 207);
			this.chkModificaImpuesto.Name = "chkModificaImpuesto";
			this.chkModificaImpuesto.Size = new System.Drawing.Size(232, 21);
			this.chkModificaImpuesto.TabIndex = 20;
			this.chkModificaImpuesto.Text = "Modificar Impuestos en Venta";
			// 
			// chkModificaPrecioV
			// 
			this.chkModificaPrecioV.BackColor = System.Drawing.Color.Transparent;
			this.chkModificaPrecioV.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkModificaPrecioV.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CambiarPrecio"));
			this.chkModificaPrecioV.Location = new System.Drawing.Point(24, 181);
			this.chkModificaPrecioV.Name = "chkModificaPrecioV";
			this.chkModificaPrecioV.Size = new System.Drawing.Size(232, 21);
			this.chkModificaPrecioV.TabIndex = 19;
			this.chkModificaPrecioV.Text = "Modificar Precios en Ventas";
			this.toolTip1.SetToolTip(this.chkModificaPrecioV, "Habilita escribir los precios en Ventas y Cotizaciones");
			// 
			// ultraCheckEditor4
			// 
			this.ultraCheckEditor4.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor4.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor4.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.IncluyeIva"));
			this.ultraCheckEditor4.Location = new System.Drawing.Point(24, 146);
			this.ultraCheckEditor4.Name = "ultraCheckEditor4";
			this.ultraCheckEditor4.Size = new System.Drawing.Size(232, 22);
			this.ultraCheckEditor4.TabIndex = 18;
			this.ultraCheckEditor4.Text = "Precios de artículos incluyen IVA";
			// 
			// optListaPrecio
			// 
			this.optListaPrecio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.TipoPrecio"));
			this.optListaPrecio.ItemAppearance = appearance16;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Precio por Cliente";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Precio por Cantidad";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Precio por Forma de Pago";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Precio por Listas";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Precios Manuales";
			this.optListaPrecio.Items.Add(valueListItem1);
			this.optListaPrecio.Items.Add(valueListItem2);
			this.optListaPrecio.Items.Add(valueListItem3);
			this.optListaPrecio.Items.Add(valueListItem4);
			this.optListaPrecio.Items.Add(valueListItem5);
			this.optListaPrecio.ItemSpacingHorizontal = 10;
			this.optListaPrecio.ItemSpacingVertical = 7;
			this.optListaPrecio.Location = new System.Drawing.Point(24, 17);
			this.optListaPrecio.Name = "optListaPrecio";
			this.optListaPrecio.Size = new System.Drawing.Size(216, 121);
			this.optListaPrecio.TabIndex = 17;
			// 
			// ultraTabPageControl7
			// 
			this.ultraTabPageControl7.Controls.Add(this.chkCuponWeb);
			this.ultraTabPageControl7.Controls.Add(this.chkEncuesta);
			this.ultraTabPageControl7.Controls.Add(this.ultraCheckEditor59);
			this.ultraTabPageControl7.Controls.Add(this.ultraTextEditor2);
			this.ultraTabPageControl7.Controls.Add(this.ultraTextEditor1);
			this.ultraTabPageControl7.Controls.Add(this.txtDirEtiquetas);
			this.ultraTabPageControl7.Controls.Add(this.groupBox11);
			this.ultraTabPageControl7.Controls.Add(this.groupBox10);
			this.ultraTabPageControl7.Controls.Add(this.groupBox6);
			this.ultraTabPageControl7.Controls.Add(this.chkConsumoFact);
			this.ultraTabPageControl7.Controls.Add(this.chkIngrEgrInterno);
			this.ultraTabPageControl7.Controls.Add(this.ultraCheckEditor37);
			this.ultraTabPageControl7.Controls.Add(this.ultraCheckEditor42);
			this.ultraTabPageControl7.Location = new System.Drawing.Point(1, 20);
			this.ultraTabPageControl7.Name = "ultraTabPageControl7";
			this.ultraTabPageControl7.Size = new System.Drawing.Size(678, 339);
			// 
			// chkEncuesta
			// 
			this.chkEncuesta.BackColor = System.Drawing.Color.Transparent;
			this.chkEncuesta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkEncuesta.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.Encuesta"));
			this.chkEncuesta.Location = new System.Drawing.Point(472, 192);
			this.chkEncuesta.Name = "chkEncuesta";
			this.chkEncuesta.Size = new System.Drawing.Size(192, 22);
			this.chkEncuesta.TabIndex = 261;
			this.chkEncuesta.Text = "Encuesta";
			// 
			// ultraCheckEditor59
			// 
			this.ultraCheckEditor59.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor59.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor59.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.GeneraSerial"));
			this.ultraCheckEditor59.Location = new System.Drawing.Point(472, 168);
			this.ultraCheckEditor59.Name = "ultraCheckEditor59";
			this.ultraCheckEditor59.Size = new System.Drawing.Size(192, 22);
			this.ultraCheckEditor59.TabIndex = 260;
			this.ultraCheckEditor59.Text = "Genera Seriales Orden de Compra";
			// 
			// ultraTextEditor2
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor2.Appearance = appearance17;
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.ADLabelJoyOP1"));
			this.ultraTextEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor2.Location = new System.Drawing.Point(8, 264);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(664, 22);
			this.ultraTextEditor2.TabIndex = 259;
			// 
			// ultraTextEditor1
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor1.Appearance = appearance18;
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.ADLabelJoyOP"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(8, 288);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(664, 22);
			this.ultraTextEditor1.TabIndex = 258;
			// 
			// txtDirEtiquetas
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDirEtiquetas.Appearance = appearance19;
			this.txtDirEtiquetas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteoF, "SeteoF.ADLabelJoy"));
			this.txtDirEtiquetas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDirEtiquetas.Location = new System.Drawing.Point(8, 312);
			this.txtDirEtiquetas.Name = "txtDirEtiquetas";
			this.txtDirEtiquetas.Size = new System.Drawing.Size(664, 22);
			this.txtDirEtiquetas.TabIndex = 257;
			// 
			// groupBox11
			// 
			this.groupBox11.Controls.Add(this.ultraCheckEditor57);
			this.groupBox11.Controls.Add(this.ultraCheckEditor56);
			this.groupBox11.Controls.Add(this.ultraCheckEditor53);
			this.groupBox11.Controls.Add(this.ultraCheckEditor54);
			this.groupBox11.Controls.Add(this.ultraCheckEditor55);
			this.groupBox11.Location = new System.Drawing.Point(480, 8);
			this.groupBox11.Name = "groupBox11";
			this.groupBox11.Size = new System.Drawing.Size(192, 152);
			this.groupBox11.TabIndex = 90;
			this.groupBox11.TabStop = false;
			this.groupBox11.Text = "Cotización";
			// 
			// ultraCheckEditor57
			// 
			this.ultraCheckEditor57.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor57.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor57.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CreaNotaEntrega"));
			this.ultraCheckEditor57.Location = new System.Drawing.Point(16, 120);
			this.ultraCheckEditor57.Name = "ultraCheckEditor57";
			this.ultraCheckEditor57.Size = new System.Drawing.Size(168, 22);
			this.ultraCheckEditor57.TabIndex = 94;
			this.ultraCheckEditor57.Text = "Nota De Entrega";
			// 
			// ultraCheckEditor56
			// 
			this.ultraCheckEditor56.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor56.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor56.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CotizaCreaProyecto"));
			this.ultraCheckEditor56.Location = new System.Drawing.Point(16, 96);
			this.ultraCheckEditor56.Name = "ultraCheckEditor56";
			this.ultraCheckEditor56.Size = new System.Drawing.Size(168, 22);
			this.ultraCheckEditor56.TabIndex = 93;
			this.ultraCheckEditor56.Text = "Crea Proyecto";
			// 
			// ultraCheckEditor53
			// 
			this.ultraCheckEditor53.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor53.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor53.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CotizaServicios"));
			this.ultraCheckEditor53.Location = new System.Drawing.Point(16, 48);
			this.ultraCheckEditor53.Name = "ultraCheckEditor53";
			this.ultraCheckEditor53.Size = new System.Drawing.Size(168, 22);
			this.ultraCheckEditor53.TabIndex = 91;
			this.ultraCheckEditor53.Text = "Servicios";
			// 
			// ultraCheckEditor54
			// 
			this.ultraCheckEditor54.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor54.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor54.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CotizaInventario"));
			this.ultraCheckEditor54.Location = new System.Drawing.Point(16, 24);
			this.ultraCheckEditor54.Name = "ultraCheckEditor54";
			this.ultraCheckEditor54.Size = new System.Drawing.Size(168, 22);
			this.ultraCheckEditor54.TabIndex = 92;
			this.ultraCheckEditor54.Text = "Inventario";
			// 
			// ultraCheckEditor55
			// 
			this.ultraCheckEditor55.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor55.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor55.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CotizaExternos"));
			this.ultraCheckEditor55.Location = new System.Drawing.Point(16, 72);
			this.ultraCheckEditor55.Name = "ultraCheckEditor55";
			this.ultraCheckEditor55.Size = new System.Drawing.Size(168, 22);
			this.ultraCheckEditor55.TabIndex = 92;
			this.ultraCheckEditor55.Text = "Cotización externa";
			// 
			// groupBox10
			// 
			this.groupBox10.Controls.Add(this.ultraCheckEditor51);
			this.groupBox10.Controls.Add(this.chkCombos);
			this.groupBox10.Controls.Add(this.ultraCheckEditor49);
			this.groupBox10.Controls.Add(this.ultraCheckEditor52);
			this.groupBox10.Controls.Add(this.ultraCheckEditor50);
			this.groupBox10.Controls.Add(this.ultraCheckEditor47);
			this.groupBox10.Controls.Add(this.ultraCheckEditor46);
			this.groupBox10.Controls.Add(this.ultraCheckEditor45);
			this.groupBox10.Controls.Add(this.ultraCheckEditor48);
			this.groupBox10.Controls.Add(this.ultraCheckEditor58);
			this.groupBox10.Location = new System.Drawing.Point(296, 8);
			this.groupBox10.Name = "groupBox10";
			this.groupBox10.Size = new System.Drawing.Size(168, 264);
			this.groupBox10.TabIndex = 35;
			this.groupBox10.TabStop = false;
			this.groupBox10.Text = "Inventario";
			// 
			// ultraCheckEditor51
			// 
			this.ultraCheckEditor51.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor51.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor51.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.InvVerPedido"));
			this.ultraCheckEditor51.Location = new System.Drawing.Point(16, 160);
			this.ultraCheckEditor51.Name = "ultraCheckEditor51";
			this.ultraCheckEditor51.Size = new System.Drawing.Size(144, 22);
			this.ultraCheckEditor51.TabIndex = 31;
			this.ultraCheckEditor51.Text = "Pedido";
			// 
			// chkCombos
			// 
			this.chkCombos.BackColor = System.Drawing.Color.Transparent;
			this.chkCombos.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkCombos.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.InvCombos"));
			this.chkCombos.Location = new System.Drawing.Point(16, 40);
			this.chkCombos.Name = "chkCombos";
			this.chkCombos.Size = new System.Drawing.Size(144, 22);
			this.chkCombos.TabIndex = 30;
			this.chkCombos.Text = "Combos";
			// 
			// ultraCheckEditor49
			// 
			this.ultraCheckEditor49.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor49.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor49.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.InvVerExSucursales"));
			this.ultraCheckEditor49.Location = new System.Drawing.Point(16, 184);
			this.ultraCheckEditor49.Name = "ultraCheckEditor49";
			this.ultraCheckEditor49.Size = new System.Drawing.Size(144, 22);
			this.ultraCheckEditor49.TabIndex = 32;
			this.ultraCheckEditor49.Text = "Existencias Sucursales";
			// 
			// ultraCheckEditor52
			// 
			this.ultraCheckEditor52.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor52.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor52.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.InvPrecioCredito"));
			this.ultraCheckEditor52.Location = new System.Drawing.Point(16, 136);
			this.ultraCheckEditor52.Name = "ultraCheckEditor52";
			this.ultraCheckEditor52.Size = new System.Drawing.Size(144, 22);
			this.ultraCheckEditor52.TabIndex = 34;
			this.ultraCheckEditor52.Text = "Precio CrediPoint";
			// 
			// ultraCheckEditor50
			// 
			this.ultraCheckEditor50.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor50.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor50.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.InvVerMercTransito"));
			this.ultraCheckEditor50.Location = new System.Drawing.Point(16, 208);
			this.ultraCheckEditor50.Name = "ultraCheckEditor50";
			this.ultraCheckEditor50.Size = new System.Drawing.Size(144, 21);
			this.ultraCheckEditor50.TabIndex = 33;
			this.ultraCheckEditor50.Text = "Mercaderia en Tránsito";
			// 
			// ultraCheckEditor47
			// 
			this.ultraCheckEditor47.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor47.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor47.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.InvPrecioDistribuidor"));
			this.ultraCheckEditor47.Location = new System.Drawing.Point(16, 64);
			this.ultraCheckEditor47.Name = "ultraCheckEditor47";
			this.ultraCheckEditor47.Size = new System.Drawing.Size(144, 22);
			this.ultraCheckEditor47.TabIndex = 27;
			this.ultraCheckEditor47.Text = "Precio Distribuidor";
			// 
			// ultraCheckEditor46
			// 
			this.ultraCheckEditor46.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor46.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor46.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.InvPrecioTarjeta"));
			this.ultraCheckEditor46.Location = new System.Drawing.Point(16, 112);
			this.ultraCheckEditor46.Name = "ultraCheckEditor46";
			this.ultraCheckEditor46.Size = new System.Drawing.Size(144, 21);
			this.ultraCheckEditor46.TabIndex = 29;
			this.ultraCheckEditor46.Text = "Precio Tarjeta";
			// 
			// ultraCheckEditor45
			// 
			this.ultraCheckEditor45.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor45.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor45.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.InvPrecioEfectivo"));
			this.ultraCheckEditor45.Location = new System.Drawing.Point(16, 88);
			this.ultraCheckEditor45.Name = "ultraCheckEditor45";
			this.ultraCheckEditor45.Size = new System.Drawing.Size(144, 22);
			this.ultraCheckEditor45.TabIndex = 28;
			this.ultraCheckEditor45.Text = "Precio Efectivo";
			// 
			// ultraCheckEditor48
			// 
			this.ultraCheckEditor48.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor48.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor48.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.InvReemplazo"));
			this.ultraCheckEditor48.Location = new System.Drawing.Point(16, 16);
			this.ultraCheckEditor48.Name = "ultraCheckEditor48";
			this.ultraCheckEditor48.Size = new System.Drawing.Size(144, 22);
			this.ultraCheckEditor48.TabIndex = 36;
			this.ultraCheckEditor48.Text = "Reemplazo";
			// 
			// ultraCheckEditor58
			// 
			this.ultraCheckEditor58.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor58.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor58.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.InventarioZoom"));
			this.ultraCheckEditor58.Location = new System.Drawing.Point(16, 232);
			this.ultraCheckEditor58.Name = "ultraCheckEditor58";
			this.ultraCheckEditor58.Size = new System.Drawing.Size(144, 22);
			this.ultraCheckEditor58.TabIndex = 96;
			this.ultraCheckEditor58.Text = "Inventario ZOOM";
			// 
			// groupBox6
			// 
			this.groupBox6.Controls.Add(this.chkCotizaReserva);
			this.groupBox6.Controls.Add(this.ultraCheckEditor27);
			this.groupBox6.Controls.Add(this.ultraCheckEditor39);
			this.groupBox6.Controls.Add(this.ultraCheckEditor29);
			this.groupBox6.Location = new System.Drawing.Point(8, 8);
			this.groupBox6.Name = "groupBox6";
			this.groupBox6.Size = new System.Drawing.Size(272, 138);
			this.groupBox6.TabIndex = 26;
			this.groupBox6.TabStop = false;
			this.groupBox6.Text = "Transformaciones";
			// 
			// chkCotizaReserva
			// 
			this.chkCotizaReserva.BackColor = System.Drawing.Color.Transparent;
			this.chkCotizaReserva.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkCotizaReserva.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CotizaEgreso"));
			this.chkCotizaReserva.Location = new System.Drawing.Point(16, 26);
			this.chkCotizaReserva.Name = "chkCotizaReserva";
			this.chkCotizaReserva.Size = new System.Drawing.Size(240, 21);
			this.chkCotizaReserva.TabIndex = 22;
			this.chkCotizaReserva.Text = "Cotización Genera Egreso / Factura";
			// 
			// ultraCheckEditor27
			// 
			this.ultraCheckEditor27.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor27.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor27.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.OrdenCIngreso"));
			this.ultraCheckEditor27.Location = new System.Drawing.Point(16, 78);
			this.ultraCheckEditor27.Name = "ultraCheckEditor27";
			this.ultraCheckEditor27.Size = new System.Drawing.Size(240, 21);
			this.ultraCheckEditor27.TabIndex = 27;
			this.ultraCheckEditor27.Text = "Orden de Compra se Transforma con Lista";
			// 
			// ultraCheckEditor39
			// 
			this.ultraCheckEditor39.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor39.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor39.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CotizaReserva"));
			this.ultraCheckEditor39.Location = new System.Drawing.Point(16, 103);
			this.ultraCheckEditor39.Name = "ultraCheckEditor39";
			this.ultraCheckEditor39.Size = new System.Drawing.Size(240, 22);
			this.ultraCheckEditor39.TabIndex = 31;
			this.ultraCheckEditor39.Text = "Cotización Genera Reservación / Factura";
			// 
			// ultraCheckEditor29
			// 
			this.ultraCheckEditor29.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor29.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor29.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.OrdenCPedido"));
			this.ultraCheckEditor29.Location = new System.Drawing.Point(16, 52);
			this.ultraCheckEditor29.Name = "ultraCheckEditor29";
			this.ultraCheckEditor29.Size = new System.Drawing.Size(240, 21);
			this.ultraCheckEditor29.TabIndex = 28;
			this.ultraCheckEditor29.Text = "Orden de Compra Genera Pedido / Compra";
			// 
			// chkConsumoFact
			// 
			this.chkConsumoFact.BackColor = System.Drawing.Color.Transparent;
			this.chkConsumoFact.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkConsumoFact.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.ConsumoFact"));
			this.chkConsumoFact.Location = new System.Drawing.Point(8, 208);
			this.chkConsumoFact.Name = "chkConsumoFact";
			this.chkConsumoFact.Size = new System.Drawing.Size(240, 22);
			this.chkConsumoFact.TabIndex = 19;
			this.chkConsumoFact.Text = "Consumo en Factura por Receta";
			// 
			// chkIngrEgrInterno
			// 
			this.chkIngrEgrInterno.BackColor = System.Drawing.Color.Transparent;
			this.chkIngrEgrInterno.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkIngrEgrInterno.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.IngEgrInterno"));
			this.chkIngrEgrInterno.Location = new System.Drawing.Point(8, 232);
			this.chkIngrEgrInterno.Name = "chkIngrEgrInterno";
			this.chkIngrEgrInterno.Size = new System.Drawing.Size(240, 21);
			this.chkIngrEgrInterno.TabIndex = 20;
			this.chkIngrEgrInterno.Text = "Ingresos y Egresos Internos de la Empresa";
			// 
			// ultraCheckEditor37
			// 
			this.ultraCheckEditor37.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor37.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor37.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.Promocion"));
			this.ultraCheckEditor37.Location = new System.Drawing.Point(8, 184);
			this.ultraCheckEditor37.Name = "ultraCheckEditor37";
			this.ultraCheckEditor37.Size = new System.Drawing.Size(240, 22);
			this.ultraCheckEditor37.TabIndex = 19;
			this.ultraCheckEditor37.Text = "Promoción";
			// 
			// ultraCheckEditor42
			// 
			this.ultraCheckEditor42.BackColor = System.Drawing.Color.Transparent;
			this.ultraCheckEditor42.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor42.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.ConsumoOrden"));
			this.ultraCheckEditor42.Location = new System.Drawing.Point(8, 160);
			this.ultraCheckEditor42.Name = "ultraCheckEditor42";
			this.ultraCheckEditor42.Size = new System.Drawing.Size(240, 22);
			this.ultraCheckEditor42.TabIndex = 22;
			this.ultraCheckEditor42.Text = "Orden P. genera consumo";
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(8, 376);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(88, 25);
			this.btAceptar.TabIndex = 8;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(104, 376);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(88, 25);
			this.btCancelar.TabIndex = 9;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl3);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl4);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl5);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl6);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl7);
			this.ultraTabControl1.Location = new System.Drawing.Point(8, 8);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.ScrollButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(680, 360);
			this.ultraTabControl1.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.ultraTabControl1.TabIndex = 10;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Contabilidad";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Ver";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Numeración";
			ultraTab4.TabPage = this.ultraTabPageControl4;
			ultraTab4.Text = "Predefinido";
			ultraTab5.TabPage = this.ultraTabPageControl5;
			ultraTab5.Text = "Decimales";
			ultraTab6.TabPage = this.ultraTabPageControl6;
			ultraTab6.Text = "Precios";
			ultraTab7.TabPage = this.ultraTabPageControl7;
			ultraTab7.Text = "Personalizado";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2,
																																															 ultraTab3,
																																															 ultraTab4,
																																															 ultraTab5,
																																															 ultraTab6,
																																															 ultraTab7});
			this.ultraTabControl1.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(678, 339);
			// 
			// chkCuponWeb
			// 
			this.chkCuponWeb.BackColor = System.Drawing.Color.Transparent;
			this.chkCuponWeb.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkCuponWeb.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeteoF, "SeteoF.CuponWeb"));
			this.chkCuponWeb.Location = new System.Drawing.Point(472, 216);
			this.chkCuponWeb.Name = "chkCuponWeb";
			this.chkCuponWeb.Size = new System.Drawing.Size(192, 22);
			this.chkCuponWeb.TabIndex = 262;
			this.chkCuponWeb.Text = "Cupon Web";
			// 
			// SeteoF
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(696, 402);
			this.Controls.Add(this.ultraTabControl1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.MaximizeBox = false;
			this.Name = "SeteoF";
			this.Text = "Propiedades de Facturación";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.SeteoF_Closing);
			this.Load += new System.EventHandler(this.SeteoF_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			this.groupBox7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.groupBox5.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			this.ultraTabPageControl2.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbSubproyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipoAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoContado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTiempoVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsNumero)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			this.groupBox9.ResumeLayout(false);
			this.groupBox8.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumLineas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDecTotal)).EndInit();
			this.ultraTabPageControl6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentajeGarantia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optListaPrecio)).EndInit();
			this.ultraTabPageControl7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDirEtiquetas)).EndInit();
			this.groupBox11.ResumeLayout(false);
			this.groupBox10.ResumeLayout(false);
			this.groupBox6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.cdsSeteoF.Update();

			this.cdsNumero.Update();

			this.Close();
		}

		private void ultraCheckEditor17_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsNumero_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsNumero.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void chkImpideSaldoNegativo_CheckedChanged(object sender, System.EventArgs e)
		{
//			if (this.chkImpideSaldoNegativo.Checked)
//				this.chkSaldoBodega.Enabled = true;
//			else
//			{
//				this.chkSaldoBodega.Enabled = false;
//				this.chkSaldoBodega.Checked = false;
//			}				
		}

		private void cmbCliente_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			this.cmbCliente.EditValue = 0;
			this.cmbCliente.Text = "";
			this.cmbCliente.Refresh();
			this.cmbCliente.Select();
		}

		private void SeteoF_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 6, 'Prop. Fact.'";
			Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
			string stSelect = "Select IsNull(idTipoAsiento, 0), IsNull(Tipo, 0) From AsientoTipo";
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);
			udsTipoAsiento.Rows.Clear();
			int iNumLinea = 0;
			while(dr.Read())
			{
				udsTipoAsiento.Rows.Add();
				udsTipoAsiento.Rows[iNumLinea]["idTipoAsiento"] = dr.GetInt32(0);
				udsTipoAsiento.Rows[iNumLinea]["TipoAsiento"] = dr.GetString(1);
				iNumLinea++;
			}
			dr.Close();
			string stExec = "Exec ClienteCargaNombre 1";
			cmbCliente.Properties.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
		}

		private void cmdListaD_Click(object sender, System.EventArgs e)
		{
			string stSelect = "Update DetCompra Set DetCompra.Codigo = Articulo.Codigo "
				+ ", DetCompra.Articulo = Articulo.Articulo "
				+ "From DetCompra Inner Join Articulo On DetCompra.idArticulo = Articulo.idArticulo "
				+ "Where IsNull(DetCompra.Codigo, '') = '' Or IsNull(DetCompra.Articulo, '') = ''";
			Funcion.EjecutaSQL(cdsSeteoF, stSelect, true);
			MessageBox.Show("Articulos Modificados", "Información");
		}

		private void chkDesplegable_CheckedChanged(object sender, System.EventArgs e)
		{
			//chkCargaArt.Enabled = !chkDesplegable.Checked;
		}

		private void btNumero_Click(object sender, System.EventArgs e)
		{
			CompraNumeroEstacion miNumero = new CompraNumeroEstacion();
			miNumero.MdiParent = this.MdiParent;
			miNumero.Show();
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsClientes"];

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
				"[Ambos] <> 0 OR [Proveedor] = 0"));
		}

		private void SeteoF_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Prop. Fact.'";
			Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void btNoAsignar_Click(object sender, System.EventArgs e)
		{
			ultraGrid1.ActiveRow.Cells["Proyecto"].Value = 0;
			ultraGrid1.ActiveRow.Cells["SubProyecto"].Value = 0;
		}

		private void btidClienteToNombre_Click(object sender, System.EventArgs e)
		{
			string stSelect = "Update Compra Set Nombre = Cliente.Nombre, Ruc = Cliente.Ruc, idTipoRuc = Cliente.idTipoRuc "
				+ "From Compra Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
				+ "Where IsNull(Compra.Nombre, '') = '' Or IsNull(Compra.Ruc, '') = '' And Compra.idCompra > 44";
			Funcion.EjecutaSQL(cdsSeteoF, stSelect, true);

			MessageBox.Show("Clientes vacios creados", "Información");
		}

		private void cmbCliente_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbCliente.EditValue = 0;
		}

		private void txtCredito_Enter(object sender, System.EventArgs e)
		{
			this.txtCredito.SelectAll();
		}

		private void txtCredito_Click(object sender, System.EventArgs e)
		{
			this.txtCredito.SelectAll();
		}

		private void txtTarjeta_Click(object sender, System.EventArgs e)
		{
			this.txtTarjeta.SelectAll();
		}

		private void txtTarjeta_Enter(object sender, System.EventArgs e)
		{
			this.txtTarjeta.SelectAll();
		}

		private void btnIVA_Click(object sender, System.EventArgs e)
		{
			frmImpuestoIVA IIVA = new frmImpuestoIVA();
			IIVA.ShowDialog();
		}

	}
}
