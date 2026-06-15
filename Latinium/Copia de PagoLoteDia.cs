using System;
using System.Data;
using System.Globalization;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de PagoLotes.
	/// </summary>
	public class PagoLoteDia : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsCompra;
		private System.ComponentModel.IContainer components;
		private C1.Data.C1DataSet cdsCliente;
		private C1.Data.C1DataView cdvCliente;
		private System.Windows.Forms.ImageList imageList;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentaCorr;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDesglosa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAsiento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkIncluyenombre;
		private Infragistics.Win.Misc.UltraButton btAsiento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdAsiento;
		private C1.Data.C1DataSet cdsAsiento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.SimpleButton btRetencion;
		private C1.Data.C1DataSet cdsCompraTabla;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbNumCuenta;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago1;
		private C1.Data.C1DataSet cdsAsignaCuenta;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCobrador1;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.Misc.UltraButton btFactura;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBanco1;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAsiento;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReferencia;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDeseado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsPagos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCliente;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbTipo;
		private C1.Data.C1DataSet cdsPago;
		private Infragistics.Win.Misc.UltraButton btPago;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsPagoLote;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuentaCorr;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTdoBanco;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor idDeposito;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor idPagoLote;
		private Infragistics.Win.Misc.UltraButton btQuitar;
		private Infragistics.Win.Misc.UltraButton btAgregar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVencimiento;
		private System.Windows.Forms.Label label3;

		int IdTipoFactura = 0;
		public PagoLoteDia()
		{
			InitializeComponent();
		}

		public PagoLoteDia(int idTipoFactura)
		{
			InitializeComponent();
			IdTipoFactura = idTipoFactura;
		}
		int iIdPagoLote = 0;
		public PagoLoteDia(int idTipoFactura, int idPagoLote)
		{
			InitializeComponent();
			IdTipoFactura = idTipoFactura;
			iIdPagoLote = idPagoLote;
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cliente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ambos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconProv");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DesconCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(PagoLoteDia));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraFormaPago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vecimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pendiente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vecimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco", -1, "cmbBanco1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente", -1, "cmbNumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pendiente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Pago", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco", -1, "cmbBanco1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador", -1, "cmbCobrador1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente", -1, "cmbNumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pendiente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PendLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PendLote1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiRet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AsiEgr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn205 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDeposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn206 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn207 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloquear");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn208 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn209 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn210 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsPagoLote = new C1.Data.C1DataSet();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.cdsPago = new C1.Data.C1DataSet();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.idDeposito = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsFormaPago = new C1.Data.C1DataSet();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.cmbCuentaCorr = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.label6 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.txtAsiento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkDesglosa = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkIncluyenombre = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btRetencion = new DevExpress.XtraEditors.SimpleButton();
			this.cmbNumCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsCompraTabla = new C1.Data.C1DataSet();
			this.cmbFormaPago1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsAsignaCuenta = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cmbCobrador1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.btFactura = new Infragistics.Win.Misc.UltraButton();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBanco1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.chkAsiento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDeseado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.udsPagos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbTipo = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btPago = new Infragistics.Win.Misc.UltraButton();
			this.btQuitar = new Infragistics.Win.Misc.UltraButton();
			this.chkCuentaCorr = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkFormaPago = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkTdoBanco = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.idPagoLote = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btAgregar = new Infragistics.Win.Misc.UltraButton();
			this.chkVencimiento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label3 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagoLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.idDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNumCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeseado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.idPagoLote)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.EnforceConstraints = false;
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// cmbCliente
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCliente.Appearance = appearance1;
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idCliente"));
			this.cmbCliente.DataSource = this.cdvCliente;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Width = 114;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 313;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn10.Header.VisiblePosition = 9;
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
																										 ultraGridColumn10});
			this.cmbCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(120, 8);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(272, 21);
			this.cmbCliente.TabIndex = 1;
			this.cmbCliente.ValueMember = "idCliente";
			// 
			// cdsPagoLote
			// 
			this.cdsPagoLote.BindingContextCtrl = this;
			this.cdsPagoLote.DataLibrary = "LibFacturacion";
			this.cdsPagoLote.DataLibraryUrl = "";
			this.cdsPagoLote.DataSetDef = "dsPagoLote";
			this.cdsPagoLote.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPagoLote.Name = "cdsPagoLote";
			this.cdsPagoLote.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPagoLote.SchemaDef = null;
			this.cdsPagoLote.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsPagoLote_PositionChanged);
			this.cdsPagoLote.BeforeFill += new C1.Data.FillEventHandler(this.cdsPagoLote_BeforeFill);
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// imageList
			// 
			this.imageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth16Bit;
			this.imageList.ImageSize = new System.Drawing.Size(18, 18);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btFiltro
			// 
			appearance2.Image = 5;
			this.btFiltro.Appearance = appearance2;
			this.btFiltro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btFiltro.ImageList = this.imageList;
			this.btFiltro.Location = new System.Drawing.Point(400, 8);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(24, 24);
			this.btFiltro.TabIndex = 7;
			this.btFiltro.Visible = false;
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// cdsPago
			// 
			this.cdsPago.BindingContextCtrl = this;
			this.cdsPago.DataLibrary = "LibFacturacion";
			this.cdsPago.DataLibraryUrl = "";
			this.cdsPago.DataSetDef = "dsPago";
			this.cdsPago.FillOnRequest = false;
			this.cdsPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPago.Name = "cdsPago";
			this.cdsPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsPago.SchemaDef = null;
			// 
			// cmbDesde
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDesde.Appearance = appearance3;
			this.cmbDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(64, 104);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 9;
			this.cmbDesde.Value = ((object)(resources.GetObject("cmbDesde.Value")));
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 104);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(40, 16);
			this.label2.TabIndex = 10;
			this.label2.Text = "Desde:";
			// 
			// idDeposito
			// 
			this.idDeposito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idDeposito"));
			this.idDeposito.Location = new System.Drawing.Point(544, 320);
			this.idDeposito.Name = "idDeposito";
			this.idDeposito.Size = new System.Drawing.Size(40, 21);
			this.idDeposito.TabIndex = 12;
			// 
			// cmbFormaPago
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance4;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idFormaPago"));
			this.cmbFormaPago.DataMember = "CompraFormaPago";
			this.cmbFormaPago.DataSource = this.cdsFormaPago;
			this.cmbFormaPago.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "CompraComprobante";
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Width = 73;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn16.Width = 37;
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn17.Width = 47;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(120, 56);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(176, 21);
			this.cmbFormaPago.TabIndex = 13;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			// 
			// cdsFormaPago
			// 
			this.cdsFormaPago.BindingContextCtrl = this;
			this.cdsFormaPago.DataLibrary = "LibFacturacion";
			this.cdsFormaPago.DataLibraryUrl = "";
			this.cdsFormaPago.DataSetDef = "dsCompraTabla";
			this.cdsFormaPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFormaPago.Name = "cdsFormaPago";
			this.cdsFormaPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFormaPago.SchemaDef = null;
			this.cdsFormaPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsFormaPago_BeforeFill);
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
			// cmbCuentaCorr
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentaCorr.Appearance = appearance5;
			this.cmbCuentaCorr.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentaCorr.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorr.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idCuentaCorriente"));
			this.cmbCuentaCorr.DataMember = "CuentaCorriente";
			this.cmbCuentaCorr.DataSource = this.cdsCuentaCorr;
			this.cmbCuentaCorr.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "CompraComprobante";
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 9;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 155;
			ultraGridColumn20.Header.VisiblePosition = 6;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Width = 88;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 69;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 8;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 3;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 4;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 10;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 40;
			ultraGridColumn29.Header.VisiblePosition = 11;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 56;
			ultraGridBand3.Columns.AddRange(new object[] {
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
			this.cmbCuentaCorr.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCuentaCorr.DisplayMember = "NumCuenta";
			this.cmbCuentaCorr.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentaCorr.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuentaCorr.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorr.Location = new System.Drawing.Point(440, 56);
			this.cmbCuentaCorr.Name = "cmbCuentaCorr";
			this.cmbCuentaCorr.Size = new System.Drawing.Size(176, 21);
			this.cmbCuentaCorr.TabIndex = 15;
			this.cmbCuentaCorr.ValueMember = "idCuentaCorriente";
			this.cmbCuentaCorr.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCuentaCorr_MouseDown);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.RowFilter = "Cobrador = 1";
			this.cdvPersonal.Sort = "Nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(312, 80);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(28, 16);
			this.label6.TabIndex = 19;
			this.label6.Text = "Nota";
			// 
			// txtNotas
			// 
			this.txtNotas.AlwaysInEditMode = true;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance6;
			this.txtNotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.Nota"));
			this.txtNotas.Location = new System.Drawing.Point(368, 80);
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(248, 21);
			this.txtNotas.TabIndex = 20;
			// 
			// txtCodigo
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance7;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.Codigo"));
			this.txtCodigo.Location = new System.Drawing.Point(72, 32);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.ReadOnly = true;
			this.txtCodigo.Size = new System.Drawing.Size(80, 21);
			this.txtCodigo.TabIndex = 21;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(16, 32);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(43, 16);
			this.label7.TabIndex = 22;
			this.label7.Text = "Código:";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(176, 32);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(45, 16);
			this.label8.TabIndex = 24;
			this.label8.Text = "Asiento:";
			// 
			// txtAsiento
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAsiento.Appearance = appearance8;
			this.txtAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.CodAsiento"));
			this.txtAsiento.Location = new System.Drawing.Point(240, 32);
			this.txtAsiento.Name = "txtAsiento";
			this.txtAsiento.ReadOnly = true;
			this.txtAsiento.Size = new System.Drawing.Size(80, 21);
			this.txtAsiento.TabIndex = 23;
			// 
			// chkDesglosa
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDesglosa.Appearance = appearance9;
			this.chkDesglosa.Checked = true;
			this.chkDesglosa.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			this.chkDesglosa.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPagoLote, "_PagoLote.DesglosarPagos"));
			this.chkDesglosa.Location = new System.Drawing.Point(336, 32);
			this.chkDesglosa.Name = "chkDesglosa";
			this.chkDesglosa.Size = new System.Drawing.Size(104, 20);
			this.chkDesglosa.TabIndex = 25;
			this.chkDesglosa.Text = "Desglosar Pago";
			// 
			// chkIncluyenombre
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkIncluyenombre.Appearance = appearance10;
			this.chkIncluyenombre.Checked = true;
			this.chkIncluyenombre.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			this.chkIncluyenombre.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPagoLote, "_PagoLote.IncluyeNombre"));
			this.chkIncluyenombre.Location = new System.Drawing.Point(456, 32);
			this.chkIncluyenombre.Name = "chkIncluyenombre";
			this.chkIncluyenombre.Size = new System.Drawing.Size(96, 20);
			this.chkIncluyenombre.TabIndex = 26;
			this.chkIncluyenombre.Text = "Incluir Nombre";
			// 
			// btAsiento
			// 
			this.btAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiento.Location = new System.Drawing.Point(328, 351);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(24, 23);
			this.btAsiento.TabIndex = 27;
			this.btAsiento.Text = "A";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			this.btAsiento.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAsiento_MouseDown);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idAsiento"));
			this.txtIdAsiento.Location = new System.Drawing.Point(504, 320);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.Size = new System.Drawing.Size(32, 21);
			this.txtIdAsiento.TabIndex = 28;
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
			this.cdsAsiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsiento_BeforeFill);
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
			// btRetencion
			// 
			this.btRetencion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btRetencion.ImageIndex = 6;
			this.btRetencion.ImageList = this.imageList;
			this.btRetencion.Location = new System.Drawing.Point(352, 351);
			this.btRetencion.Name = "btRetencion";
			this.btRetencion.Size = new System.Drawing.Size(24, 23);
			this.btRetencion.TabIndex = 29;
			this.btRetencion.ToolTip = "Retención en la fuente";
			this.btRetencion.Click += new System.EventHandler(this.btRetencion_Click);
			this.btRetencion.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btRetencion_MouseDown);
			// 
			// cmbNumCuenta
			// 
			this.cmbNumCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNumCuenta.DataMember = "CuentaCorriente";
			this.cmbNumCuenta.DataSource = this.cdsCuentaCorr;
			this.cmbNumCuenta.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn30.Header.VisiblePosition = 6;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 9;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 47;
			ultraGridColumn32.Header.VisiblePosition = 4;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 0;
			ultraGridColumn34.Width = 141;
			ultraGridColumn35.Header.VisiblePosition = 7;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 8;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 2;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 3;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 10;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 42;
			ultraGridColumn41.Header.VisiblePosition = 11;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 54;
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
																										 ultraGridColumn41});
			this.cmbNumCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbNumCuenta.DisplayMember = "NumCuenta";
			this.cmbNumCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNumCuenta.Location = new System.Drawing.Point(184, 224);
			this.cmbNumCuenta.Name = "cmbNumCuenta";
			this.cmbNumCuenta.Size = new System.Drawing.Size(160, 64);
			this.cmbNumCuenta.TabIndex = 30;
			this.cmbNumCuenta.ValueMember = "idCuentaCorriente";
			this.cmbNumCuenta.Visible = false;
			// 
			// cdsCompraTabla
			// 
			this.cdsCompraTabla.BindingContextCtrl = this;
			this.cdsCompraTabla.DataLibrary = "LibFacturacion";
			this.cdsCompraTabla.DataLibraryUrl = "";
			this.cdsCompraTabla.DataSetDef = "dsCompraTabla";
			this.cdsCompraTabla.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraTabla.Name = "cdsCompraTabla";
			this.cdsCompraTabla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraTabla.SchemaDef = null;
			this.cdsCompraTabla.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompraTabla_BeforeFill);
			// 
			// cmbFormaPago1
			// 
			this.cmbFormaPago1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago1.DataMember = "CompraFormaPago";
			this.cmbFormaPago1.DataSource = this.cdsFormaPago;
			ultraGridBand5.AddButtonCaption = "CompraComprobante";
			ultraGridColumn42.Header.VisiblePosition = 3;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 4;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 0;
			ultraGridColumn44.Width = 118;
			ultraGridColumn45.Header.VisiblePosition = 2;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 1;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 5;
			ultraGridColumn48.Header.VisiblePosition = 6;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48});
			this.cmbFormaPago1.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbFormaPago1.DisplayMember = "FormaPago";
			this.cmbFormaPago1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago1.Location = new System.Drawing.Point(424, 168);
			this.cmbFormaPago1.Name = "cmbFormaPago1";
			this.cmbFormaPago1.Size = new System.Drawing.Size(144, 80);
			this.cmbFormaPago1.TabIndex = 31;
			this.cmbFormaPago1.ValueMember = "idFormaPago";
			this.cmbFormaPago1.Visible = false;
			// 
			// cdsAsignaCuenta
			// 
			this.cdsAsignaCuenta.BindingContextCtrl = this;
			this.cdsAsignaCuenta.DataLibrary = "LibContabilidad";
			this.cdsAsignaCuenta.DataLibraryUrl = "";
			this.cdsAsignaCuenta.DataSetDef = "dsAsignaCuenta";
			this.cdsAsignaCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsignaCuenta.Name = "cdsAsignaCuenta";
			this.cdsAsignaCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsignaCuenta.SchemaDef = null;
			this.cdsAsignaCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsAsignaCuenta_BeforeFill);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_PagoLote";
			this.barraDato1.DataNombreId = "idPagoLote";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsPagoLote;
			this.barraDato1.DataTabla = "PagoLote";
			this.barraDato1.DataTablaHija = "Pago";
			this.barraDato1.Location = new System.Drawing.Point(16, 351);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 32;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cmbCobrador1
			// 
			this.cmbCobrador1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobrador1.DataSource = this.cdvPersonal;
			this.cmbCobrador1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn49.Header.VisiblePosition = 21;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 59;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 52;
			ultraGridColumn51.Header.VisiblePosition = 19;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 20;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 18;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 22;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 36;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 25;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 0;
			ultraGridColumn57.Width = 8;
			ultraGridColumn58.Header.VisiblePosition = 3;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 35;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 24;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 13;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 32;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 7;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 6;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 2;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 29;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 28;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 23;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 31;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 10;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 33;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 14;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 15;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 17;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 26;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 5;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 37;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 38;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 27;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 30;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 12;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 9;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 11;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 16;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 1;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 34;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 4;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 8;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 39;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Width = 8;
			ultraGridColumn90.Header.VisiblePosition = 40;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Width = 16;
			ultraGridColumn91.Header.VisiblePosition = 41;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 22;
			ultraGridColumn92.Header.VisiblePosition = 42;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn92.Width = 24;
			ultraGridColumn93.Header.VisiblePosition = 43;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn93.Width = 22;
			ultraGridColumn94.Header.VisiblePosition = 44;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Width = 44;
			ultraGridColumn95.Header.VisiblePosition = 45;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn95.Width = 24;
			ultraGridColumn96.Header.VisiblePosition = 46;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn96.Width = 9;
			ultraGridColumn97.Header.VisiblePosition = 47;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 19;
			ultraGridColumn98.Header.VisiblePosition = 48;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn98.Width = 25;
			ultraGridColumn99.Header.VisiblePosition = 49;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn99.Width = 42;
			ultraGridColumn100.Header.VisiblePosition = 50;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn100.Width = 62;
			ultraGridColumn101.Header.VisiblePosition = 51;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn101.Width = 35;
			ultraGridColumn102.Header.VisiblePosition = 52;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn102.Width = 32;
			ultraGridColumn103.Header.VisiblePosition = 53;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn103.Width = 42;
			ultraGridColumn104.Header.VisiblePosition = 54;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn104.Width = 17;
			ultraGridColumn105.Header.VisiblePosition = 55;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn105.Width = 19;
			ultraGridColumn106.Header.VisiblePosition = 56;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn106.Width = 33;
			ultraGridColumn107.Header.VisiblePosition = 57;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn107.Width = 44;
			ultraGridColumn108.Header.VisiblePosition = 58;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn108.Width = 33;
			ultraGridColumn109.Header.VisiblePosition = 60;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn109.Width = 34;
			ultraGridColumn110.Header.VisiblePosition = 61;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn110.Width = 49;
			ultraGridColumn111.Header.VisiblePosition = 62;
			ultraGridColumn111.Width = 8;
			ultraGridColumn112.Header.VisiblePosition = 63;
			ultraGridColumn112.Width = 17;
			ultraGridColumn113.Header.VisiblePosition = 64;
			ultraGridColumn113.Width = 17;
			ultraGridColumn114.Header.VisiblePosition = 65;
			ultraGridColumn114.Width = 8;
			ultraGridColumn115.Header.VisiblePosition = 66;
			ultraGridColumn115.Width = 17;
			ultraGridColumn116.Header.VisiblePosition = 67;
			ultraGridColumn116.Width = 8;
			ultraGridColumn117.Header.VisiblePosition = 68;
			ultraGridColumn117.Width = 8;
			ultraGridColumn118.Header.VisiblePosition = 69;
			ultraGridColumn118.Width = 8;
			ultraGridColumn119.Header.VisiblePosition = 70;
			ultraGridColumn119.Width = 8;
			ultraGridColumn120.Header.VisiblePosition = 71;
			ultraGridColumn120.Width = 8;
			ultraGridColumn121.Header.VisiblePosition = 72;
			ultraGridColumn121.Width = 8;
			ultraGridColumn122.Header.VisiblePosition = 73;
			ultraGridColumn122.Width = 8;
			ultraGridColumn123.Header.VisiblePosition = 74;
			ultraGridColumn123.Width = 8;
			ultraGridColumn124.Header.VisiblePosition = 75;
			ultraGridColumn124.Width = 8;
			ultraGridColumn125.Header.VisiblePosition = 76;
			ultraGridColumn125.Width = 8;
			ultraGridColumn126.Header.VisiblePosition = 77;
			ultraGridColumn126.Width = 8;
			ultraGridColumn127.Header.VisiblePosition = 78;
			ultraGridColumn127.Width = 8;
			ultraGridColumn128.Header.VisiblePosition = 79;
			ultraGridColumn128.Width = 8;
			ultraGridColumn129.Header.VisiblePosition = 80;
			ultraGridColumn129.Width = 8;
			ultraGridColumn130.Header.VisiblePosition = 81;
			ultraGridColumn130.Width = 8;
			ultraGridColumn131.Header.VisiblePosition = 82;
			ultraGridColumn131.Width = 8;
			ultraGridColumn132.Header.VisiblePosition = 83;
			ultraGridColumn132.Width = 8;
			ultraGridColumn133.Header.VisiblePosition = 84;
			ultraGridColumn133.Width = 8;
			ultraGridColumn134.Header.VisiblePosition = 85;
			ultraGridColumn134.Width = 8;
			ultraGridColumn135.Header.VisiblePosition = 86;
			ultraGridColumn135.Width = 8;
			ultraGridColumn136.Header.VisiblePosition = 87;
			ultraGridColumn136.Width = 8;
			ultraGridColumn137.Header.VisiblePosition = 88;
			ultraGridColumn137.Width = 8;
			ultraGridColumn138.Header.VisiblePosition = 89;
			ultraGridColumn138.Width = 8;
			ultraGridColumn139.Header.VisiblePosition = 90;
			ultraGridColumn139.Width = 8;
			ultraGridColumn140.Header.VisiblePosition = 91;
			ultraGridColumn140.Width = 8;
			ultraGridColumn141.Header.VisiblePosition = 92;
			ultraGridColumn141.Width = 8;
			ultraGridColumn142.Header.VisiblePosition = 93;
			ultraGridColumn142.Width = 8;
			ultraGridColumn143.Header.VisiblePosition = 94;
			ultraGridColumn143.Width = 17;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn143});
			this.cmbCobrador1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCobrador1.DisplayMember = "Nombre";
			this.cmbCobrador1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador1.Location = new System.Drawing.Point(496, 176);
			this.cmbCobrador1.Name = "cmbCobrador1";
			this.cmbCobrador1.Size = new System.Drawing.Size(136, 64);
			this.cmbCobrador1.TabIndex = 34;
			this.cmbCobrador1.ValueMember = "idPersonal";
			this.cmbCobrador1.Visible = false;
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
			// btFactura
			// 
			this.btFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btFactura.Location = new System.Drawing.Point(376, 351);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(24, 23);
			this.btFactura.TabIndex = 35;
			this.btFactura.Text = "P";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// cmbBanco
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance11;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idBanco"));
			this.cmbBanco.DataMember = "Banco";
			this.cmbBanco.DataSource = this.cdsCuentaCorr;
			this.cmbBanco.DisplayLayout.AutoFitColumns = true;
			ultraGridBand7.AddButtonCaption = "CompraFormaPago";
			ultraGridColumn144.Header.VisiblePosition = 4;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn145.Header.VisiblePosition = 0;
			ultraGridColumn145.Width = 69;
			ultraGridColumn146.Header.VisiblePosition = 1;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn146.Width = 81;
			ultraGridColumn147.Header.VisiblePosition = 3;
			ultraGridColumn147.Hidden = true;
			ultraGridColumn147.Width = 83;
			ultraGridColumn148.Header.VisiblePosition = 2;
			ultraGridColumn148.Hidden = true;
			ultraGridColumn149.Header.VisiblePosition = 5;
			ultraGridColumn149.Width = 43;
			ultraGridColumn150.Header.VisiblePosition = 6;
			ultraGridColumn150.Width = 45;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn144,
																										 ultraGridColumn145,
																										 ultraGridColumn146,
																										 ultraGridColumn147,
																										 ultraGridColumn148,
																										 ultraGridColumn149,
																										 ultraGridColumn150});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(120, 80);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(176, 21);
			this.cmbBanco.TabIndex = 36;
			this.cmbBanco.ValueMember = "idBanco";
			// 
			// cmbBanco1
			// 
			this.cmbBanco1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco1.DataMember = "Banco";
			this.cmbBanco1.DataSource = this.cdsCuentaCorr;
			ultraGridColumn151.Header.VisiblePosition = 3;
			ultraGridColumn151.Hidden = true;
			ultraGridColumn152.Header.VisiblePosition = 0;
			ultraGridColumn152.Width = 142;
			ultraGridColumn153.Header.VisiblePosition = 1;
			ultraGridColumn153.Hidden = true;
			ultraGridColumn154.Header.VisiblePosition = 4;
			ultraGridColumn154.Hidden = true;
			ultraGridColumn155.Header.VisiblePosition = 2;
			ultraGridColumn155.Hidden = true;
			ultraGridColumn156.Header.VisiblePosition = 5;
			ultraGridColumn157.Header.VisiblePosition = 6;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn151,
																										 ultraGridColumn152,
																										 ultraGridColumn153,
																										 ultraGridColumn154,
																										 ultraGridColumn155,
																										 ultraGridColumn156,
																										 ultraGridColumn157});
			this.cmbBanco1.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbBanco1.DisplayMember = "Nombre";
			this.cmbBanco1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco1.Location = new System.Drawing.Point(472, 224);
			this.cmbBanco1.Name = "cmbBanco1";
			this.cmbBanco1.Size = new System.Drawing.Size(160, 80);
			this.cmbBanco1.TabIndex = 37;
			this.cmbBanco1.ValueMember = "idBanco";
			this.cmbBanco1.Visible = false;
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
			// chkAsiento
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAsiento.Appearance = appearance12;
			this.chkAsiento.Checked = true;
			this.chkAsiento.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			this.chkAsiento.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsPagoLote, "_PagoLote.AsiEgr"));
			this.chkAsiento.Location = new System.Drawing.Point(560, 32);
			this.chkAsiento.Name = "chkAsiento";
			this.chkAsiento.Size = new System.Drawing.Size(72, 20);
			this.chkAsiento.TabIndex = 38;
			this.chkAsiento.Text = "Asiento";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(456, 8);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(35, 16);
			this.label9.TabIndex = 39;
			this.label9.Text = "Refer.";
			// 
			// txtReferencia
			// 
			this.txtReferencia.AlwaysInEditMode = true;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtReferencia.Appearance = appearance13;
			this.txtReferencia.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.Referencia"));
			this.txtReferencia.Location = new System.Drawing.Point(512, 8);
			this.txtReferencia.Name = "txtReferencia";
			this.txtReferencia.Size = new System.Drawing.Size(104, 21);
			this.txtReferencia.TabIndex = 40;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance14;
			this.txtTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.Total"));
			this.txtTotal.FormatString = "#,###0.00";
			this.txtTotal.Location = new System.Drawing.Point(430, 351);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(100, 21);
			this.txtTotal.TabIndex = 42;
			this.txtTotal.DoubleClick += new System.EventHandler(this.txtTotal_DoubleClick);
			// 
			// txtDeseado
			// 
			this.txtDeseado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDeseado.Appearance = appearance15;
			this.txtDeseado.FormatString = "#,###0.00";
			this.txtDeseado.Location = new System.Drawing.Point(534, 351);
			this.txtDeseado.Name = "txtDeseado";
			this.txtDeseado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDeseado.PromptChar = ' ';
			this.txtDeseado.Size = new System.Drawing.Size(100, 21);
			this.txtDeseado.TabIndex = 43;
			// 
			// udsPagos
			// 
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn3.DataType = typeof(System.Double);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.Double);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(System.Byte);
			this.udsPagos.Band.Columns.AddRange(new object[] {
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
																												 ultraDataColumn15});
			// 
			// chkCliente
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCliente.Appearance = appearance16;
			this.chkCliente.Checked = true;
			this.chkCliente.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCliente.Location = new System.Drawing.Point(16, 8);
			this.chkCliente.Name = "chkCliente";
			this.chkCliente.Size = new System.Drawing.Size(96, 20);
			this.chkCliente.TabIndex = 44;
			this.chkCliente.Text = "Todo Cliente";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(192, 104);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 46;
			this.label1.Text = "Hasta:";
			// 
			// cmbHasta
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbHasta.Appearance = appearance17;
			this.cmbHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbHasta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.Fecha"));
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(240, 104);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 45;
			this.cmbHasta.Value = ((object)(resources.GetObject("cmbHasta.Value")));
			// 
			// cmbTipo
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance18;
			this.cmbTipo.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Cobros Realizados";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Cobros Pendientes";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Todo Cobro";
			this.cmbTipo.Items.Add(valueListItem1);
			this.cmbTipo.Items.Add(valueListItem2);
			this.cmbTipo.Items.Add(valueListItem3);
			this.cmbTipo.Location = new System.Drawing.Point(368, 104);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(128, 21);
			this.cmbTipo.TabIndex = 47;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.udsPagos;
			appearance19.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance19;
			ultraGridBand9.AddButtonCaption = "Pago";
			ultraGridColumn158.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn158.Header.VisiblePosition = 4;
			ultraGridColumn159.Format = "dd/MMM/yyyy";
			ultraGridColumn159.Header.VisiblePosition = 8;
			ultraGridColumn159.Width = 80;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn160.CellAppearance = appearance20;
			ultraGridColumn160.Format = "#,##0.00";
			ultraGridColumn160.Header.VisiblePosition = 3;
			ultraGridColumn160.Width = 66;
			ultraGridColumn161.Header.VisiblePosition = 9;
			ultraGridColumn161.Width = 66;
			ultraGridColumn162.Format = "dd/MMM/yyyy";
			ultraGridColumn162.Header.VisiblePosition = 7;
			ultraGridColumn162.Width = 73;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn163.CellAppearance = appearance21;
			ultraGridColumn163.Formula = "#,##0.00";
			ultraGridColumn163.Header.VisiblePosition = 10;
			ultraGridColumn163.Width = 62;
			ultraGridColumn164.Header.Caption = "FormaPago";
			ultraGridColumn164.Header.VisiblePosition = 11;
			ultraGridColumn164.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn165.Header.Caption = "Banco";
			ultraGridColumn165.Header.VisiblePosition = 2;
			ultraGridColumn165.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn166.Header.VisiblePosition = 13;
			ultraGridColumn166.Hidden = true;
			ultraGridColumn167.Header.Caption = "Cuenta Corr.";
			ultraGridColumn167.Header.VisiblePosition = 6;
			ultraGridColumn167.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn168.Header.VisiblePosition = 12;
			ultraGridColumn169.Header.VisiblePosition = 5;
			ultraGridColumn170.Header.VisiblePosition = 1;
			ultraGridColumn170.Width = 72;
			ultraGridColumn171.Header.VisiblePosition = 0;
			ultraGridColumn172.Header.VisiblePosition = 14;
			ultraGridBand9.Columns.AddRange(new object[] {
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
																										 ultraGridColumn172});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance23.FontData.BoldAsString = "True";
			appearance23.FontData.Name = "Arial";
			appearance23.FontData.SizeInPoints = 10F;
			appearance23.ForeColor = System.Drawing.Color.White;
			appearance23.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance23;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 152);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(344, 184);
			this.ultraGrid1.TabIndex = 48;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// btPago
			// 
			appearance26.Image = 5;
			this.btPago.Appearance = appearance26;
			this.btPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPago.ImageList = this.imageList;
			this.btPago.Location = new System.Drawing.Point(504, 104);
			this.btPago.Name = "btPago";
			this.btPago.Size = new System.Drawing.Size(24, 24);
			this.btPago.TabIndex = 49;
			this.btPago.Click += new System.EventHandler(this.btPago_Click);
			// 
			// btQuitar
			// 
			this.btQuitar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btQuitar.Enabled = false;
			this.btQuitar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btQuitar.ImageList = this.imageList;
			this.btQuitar.Location = new System.Drawing.Point(368, 216);
			this.btQuitar.Name = "btQuitar";
			this.btQuitar.Size = new System.Drawing.Size(24, 23);
			this.btQuitar.TabIndex = 5;
			this.btQuitar.Text = "X";
			this.btQuitar.Click += new System.EventHandler(this.btQuitar_Click);
			// 
			// chkCuentaCorr
			// 
			this.chkCuentaCorr.Checked = true;
			this.chkCuentaCorr.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCuentaCorr.Location = new System.Drawing.Point(312, 56);
			this.chkCuentaCorr.Name = "chkCuentaCorr";
			this.chkCuentaCorr.TabIndex = 50;
			this.chkCuentaCorr.Text = "Toda Cuenta Corr.";
			// 
			// chkFormaPago
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFormaPago.Appearance = appearance27;
			this.chkFormaPago.Checked = true;
			this.chkFormaPago.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkFormaPago.Location = new System.Drawing.Point(16, 56);
			this.chkFormaPago.Name = "chkFormaPago";
			this.chkFormaPago.Size = new System.Drawing.Size(96, 20);
			this.chkFormaPago.TabIndex = 51;
			this.chkFormaPago.Text = "Toda Fr. Pago";
			// 
			// chkTdoBanco
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTdoBanco.Appearance = appearance28;
			this.chkTdoBanco.Checked = true;
			this.chkTdoBanco.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTdoBanco.Location = new System.Drawing.Point(16, 80);
			this.chkTdoBanco.Name = "chkTdoBanco";
			this.chkTdoBanco.Size = new System.Drawing.Size(96, 20);
			this.chkTdoBanco.TabIndex = 52;
			this.chkTdoBanco.Text = "Todo Banco";
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "Pago";
			this.ultraGrid2.DataSource = this.cdsPago;
			appearance29.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance29;
			ultraGridBand10.AddButtonCaption = "_Pago";
			ultraGridColumn173.Header.VisiblePosition = 21;
			ultraGridColumn173.Hidden = true;
			ultraGridColumn174.Header.VisiblePosition = 16;
			ultraGridColumn174.Hidden = true;
			ultraGridColumn175.Header.VisiblePosition = 20;
			ultraGridColumn175.Hidden = true;
			ultraGridColumn176.Header.VisiblePosition = 18;
			ultraGridColumn176.Hidden = true;
			ultraGridColumn177.Header.Caption = "Forma Pago";
			ultraGridColumn177.Header.VisiblePosition = 8;
			ultraGridColumn177.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn177.Width = 107;
			ultraGridColumn178.Header.Caption = "Banco";
			ultraGridColumn178.Header.VisiblePosition = 11;
			ultraGridColumn178.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn178.Width = 106;
			ultraGridColumn179.Header.Caption = "Cobrador";
			ultraGridColumn179.Header.VisiblePosition = 10;
			ultraGridColumn179.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn179.Width = 87;
			ultraGridColumn180.Header.Caption = "Cuenta";
			ultraGridColumn180.Header.VisiblePosition = 7;
			ultraGridColumn180.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn180.Width = 83;
			ultraGridColumn181.Header.VisiblePosition = 27;
			ultraGridColumn181.Hidden = true;
			ultraGridColumn182.Header.VisiblePosition = 9;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn183.CellAppearance = appearance30;
			ultraGridColumn183.Format = "#,##0.00";
			ultraGridColumn183.Header.VisiblePosition = 0;
			ultraGridColumn183.Width = 60;
			ultraGridColumn184.Header.VisiblePosition = 22;
			ultraGridColumn184.Hidden = true;
			ultraGridColumn185.Header.VisiblePosition = 4;
			ultraGridColumn185.Hidden = true;
			ultraGridColumn186.Header.VisiblePosition = 24;
			ultraGridColumn186.Hidden = true;
			ultraGridColumn187.Header.VisiblePosition = 1;
			ultraGridColumn187.Width = 87;
			ultraGridColumn188.Header.Caption = "Recibo";
			ultraGridColumn188.Header.VisiblePosition = 2;
			ultraGridColumn189.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn189.Header.VisiblePosition = 6;
			ultraGridColumn190.Header.VisiblePosition = 15;
			ultraGridColumn190.Hidden = true;
			ultraGridColumn191.Header.VisiblePosition = 3;
			ultraGridColumn191.Hidden = true;
			ultraGridColumn192.Header.VisiblePosition = 25;
			ultraGridColumn192.Hidden = true;
			ultraGridColumn193.Header.VisiblePosition = 12;
			ultraGridColumn194.Header.VisiblePosition = 26;
			ultraGridColumn194.Hidden = true;
			ultraGridColumn195.Header.VisiblePosition = 17;
			ultraGridColumn195.Hidden = true;
			ultraGridColumn196.Header.VisiblePosition = 23;
			ultraGridColumn196.Hidden = true;
			ultraGridColumn197.Header.VisiblePosition = 19;
			ultraGridColumn197.Hidden = true;
			ultraGridColumn198.Header.VisiblePosition = 28;
			ultraGridColumn198.Hidden = true;
			ultraGridColumn198.Width = 56;
			ultraGridColumn199.Header.VisiblePosition = 29;
			ultraGridColumn199.Hidden = true;
			ultraGridColumn200.Header.VisiblePosition = 30;
			ultraGridColumn200.Hidden = true;
			ultraGridColumn201.Header.VisiblePosition = 31;
			ultraGridColumn201.Hidden = true;
			ultraGridColumn202.Header.VisiblePosition = 32;
			ultraGridColumn202.Hidden = true;
			ultraGridColumn203.Header.VisiblePosition = 5;
			ultraGridColumn204.Header.VisiblePosition = 13;
			ultraGridColumn205.Header.VisiblePosition = 33;
			ultraGridColumn205.Hidden = true;
			ultraGridColumn206.Header.VisiblePosition = 14;
			ultraGridColumn206.Hidden = true;
			ultraGridColumn207.Header.VisiblePosition = 34;
			ultraGridColumn208.Header.VisiblePosition = 35;
			ultraGridColumn209.Header.VisiblePosition = 36;
			ultraGridColumn210.Header.VisiblePosition = 37;
			ultraGridBand10.Columns.AddRange(new object[] {
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
																											ultraGridColumn203,
																											ultraGridColumn204,
																											ultraGridColumn205,
																											ultraGridColumn206,
																											ultraGridColumn207,
																											ultraGridColumn208,
																											ultraGridColumn209,
																											ultraGridColumn210});
			ultraGridBand10.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance31.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance32.FontData.BoldAsString = "True";
			appearance32.FontData.Name = "Arial";
			appearance32.FontData.SizeInPoints = 10F;
			appearance32.ForeColor = System.Drawing.Color.White;
			appearance32.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance32;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance34;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(400, 152);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(222, 184);
			this.ultraGrid2.TabIndex = 53;
			// 
			// idPagoLote
			// 
			this.idPagoLote.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPagoLote, "_PagoLote.idPagoLote"));
			this.idPagoLote.Location = new System.Drawing.Point(456, 320);
			this.idPagoLote.Name = "idPagoLote";
			this.idPagoLote.Size = new System.Drawing.Size(40, 21);
			this.idPagoLote.TabIndex = 55;
			// 
			// btAgregar
			// 
			this.btAgregar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAgregar.Enabled = false;
			this.btAgregar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btAgregar.ImageList = this.imageList;
			this.btAgregar.Location = new System.Drawing.Point(368, 184);
			this.btAgregar.Name = "btAgregar";
			this.btAgregar.Size = new System.Drawing.Size(24, 23);
			this.btAgregar.TabIndex = 56;
			this.btAgregar.Text = ">";
			this.btAgregar.Click += new System.EventHandler(this.btAgregar_Click);
			// 
			// chkVencimiento
			// 
			this.chkVencimiento.Checked = true;
			this.chkVencimiento.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkVencimiento.Location = new System.Drawing.Point(536, 104);
			this.chkVencimiento.Name = "chkVencimiento";
			this.chkVencimiento.Size = new System.Drawing.Size(88, 20);
			this.chkVencimiento.TabIndex = 57;
			this.chkVencimiento.Text = "Vencimiento";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 128);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(40, 16);
			this.label3.TabIndex = 59;
			this.label3.Text = "Detalle";
			// 
			// PagoLoteDia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(646, 389);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.chkVencimiento);
			this.Controls.Add(this.btAgregar);
			this.Controls.Add(this.idPagoLote);
			this.Controls.Add(this.chkTdoBanco);
			this.Controls.Add(this.chkFormaPago);
			this.Controls.Add(this.chkCuentaCorr);
			this.Controls.Add(this.cmbNumCuenta);
			this.Controls.Add(this.btPago);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.chkCliente);
			this.Controls.Add(this.txtDeseado);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtReferencia);
			this.Controls.Add(this.chkAsiento);
			this.Controls.Add(this.cmbBanco1);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.cmbCobrador1);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbFormaPago1);
			this.Controls.Add(this.btRetencion);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.chkIncluyenombre);
			this.Controls.Add(this.chkDesglosa);
			this.Controls.Add(this.txtAsiento);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.cmbCuentaCorr);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.idDeposito);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.btQuitar);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.ultraGrid2);
			this.Name = "PagoLoteDia";
			this.Text = "Pagos por Lotes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.PagoLoteDia_Closing);
			this.Load += new System.EventHandler(this.PagoLotes_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPagoLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.idDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNumCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTabla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsignaCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeseado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.idPagoLote)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void BorraErrores()
		{
			this.errorProvider.SetError(this.cmbBanco, "");
			this.errorProvider.SetError(cmbHasta, "");
			this.errorProvider.SetError(this.ultraGrid1, "");
			this.errorProvider.SetError(this.ultraGrid2, "");
			this.errorProvider.SetError(this.cmbCliente, "");
			this.errorProvider.SetError(this.cmbFormaPago, "");
			this.errorProvider.SetError(this.cmbCuentaCorr, "");
		}

		private bool VerificaAgregar()
		{
			bool bOk = true;
			this.BorraErrores();
			if (cmbHasta.Value == System.DBNull.Value)
			{
				this.errorProvider.SetError(cmbHasta, "Ingrese Fecha");
				bOk = false;
			}
			if (ultraGrid1.ActiveRow == null)
			{
				this.errorProvider.SetError(this.ultraGrid1, "Seleccione pago a Depositar€");
				bOk = false;
			}
			if ((int) this.cmbFormaPago.Value == 0)
			{
				this.errorProvider.SetError(this.cmbFormaPago, "Ingrese Forma de Pago");
				bOk = false;
			}
			if(cmbCuentaCorr.Value == System.DBNull.Value || (int) cmbCuentaCorr.Value == 0)
			{
				errorProvider.SetError(this.cmbCuentaCorr, "Ingrese Cuenta corriente");
				bOk = false;
			}
			if(chkTdoBanco.Checked && (cmbBanco.Value == System.DBNull.Value || (int) this.cmbBanco.Value == 0))
			{
				this.errorProvider.SetError(this.cmbBanco, "Ingrese Banco");
				bOk = false;
			}
			return bOk;
		}

		private void btQuitar_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid2.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la fila que desea eliminar", "Información");
				return;
			}

			if (ultraGrid2.Selected.Rows.Count > 0)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Selected.Rows)
				{
					dr.Cells["idDeposito"].Value = 0;
				}
			}
			else
			{
				int IdPago = (int) ultraGrid2.ActiveRow.Cells["idPago"].Value;

				C1.Data.C1DataRow drPago = cdsPago.TableViews["Pago"].Rows.Find(IdPago);
				drPago["idDeposito"] = 0;
			}
			cdsPago.Update();
			ultraGrid2.Refresh();
			txtTotal_DoubleClick(this, e);
			barraDatoSQL1_Refresca(this, e);
			txtTotal_DoubleClick(this, e);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar Pagos", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show("Desea Borrar este Deposito?\nEsta accion no se puede revertir.",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, 
				MessageBoxDefaultButton.Button2)) return;

			string stSelect;
			if (txtIdAsiento.Value != null && int.Parse(txtIdAsiento.Value.ToString()) > 0)
			{
				stSelect = "SELECT COUNT(*) FROM Asiento WHERE idAsiento = "+
					this.txtIdAsiento.Value.ToString().Trim();
				int iCuenta = 0;
				iCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelect);

				if (iCuenta > 0)
				{
					if (DialogResult.No == MessageBox.Show("Existe un Asiento creado con este deposito. No puede Borrarlo", 
						"Información", MessageBoxButtons.OK, MessageBoxIcon.Asterisk))
						return;
				}
			}

			int IdDeposito = int.Parse(this.idDeposito.Value.ToString());
			stSelect = "Update Pago Set idDeposito = 0 Where idDeposito = " 
				+ IdDeposito.ToString();
			Funcion.EjecutaSQL(cdsCompra, stSelect);
			barraDato1.BorraRegistro();
			MessageBox.Show("Registro Borrado",
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using(BuscaPagoLotes miPagoLote = new BuscaPagoLotes(IdTipoFactura))
			{
				miPagoLote.txtDeposito.Text = "d";
				if (DialogResult.OK == miPagoLote.ShowDialog())
				{
					try
					{
						int IdBusca = int.Parse(miPagoLote.idBusca.Text);
						barraDato1.IdRegistro = IdBusca;
						barraDato1.ProximoId(5);
					}
					catch{};
				}
			}
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}


		private void Habilita(bool bActiva)
		{
			btAgregar.Enabled = bActiva;
			btQuitar.Enabled = bActiva;
		}
		private void ValoresIniciales()
		{
			drPagoLote = cdsPagoLote.TableViews["PagoLote"].Rows.Find
				(int.Parse(idPagoLote.Value.ToString()));
			drPagoLote["Fecha"] = DateTime.Today;
			drPagoLote["idTipoFactura"] = IdTipoFactura;
			string stCuenta = "Select Count(*) From PagoLote";
			int iCuenta = Funcion.iEscalarSQL(cdsCliente, stCuenta, true);
			drPagoLote["IdDeposito"] = iCuenta;
			this.Habilita(true);
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear Pagos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDato1.stFiltroExt = " And idTipoFactura = " 
				+ IdTipoFactura.ToString();// + " And IdDeposito > 0";

			this.barraDato1.CrearRegistro();
			this.drPagoLote["idTipoFactura"] = IdTipoFactura;
			//this.drPagoLote["idC"] = IdTipoFactura;
			this.barraDato1.GrabaRegistro();
			this.barraDato1.PosUltima();
			this.barraDato1.EditarRegistro();
			this.barraDato1.bNuevo = true;
			ValoresIniciales();
			barraDato1.stFiltroExt = " And idTipoFactura = " 
				+ IdTipoFactura.ToString()+ " And IdDeposito > 0";
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			cdsPago.RejectChanges();
			barraDato1.DeshacerRegistro();
			Habilita(false);
			BorraErrores();
			ultraGrid1.Enabled = true;
			if (DialogResult.No == MessageBox.Show("¿Desea Eliminar los Depositos realizados en ésta transacción?",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			int IdDeposito = -1;
			if (idDeposito.Text.Length > 0)
				IdDeposito = int.Parse(idDeposito.Text);

			string stSelect = "Update Pago Set idDeposito = 0 Where idDeposito = " + IdDeposito.ToString();
			Funcion.EjecutaSQL(cdsCliente, stSelect, false);
			barraDatoSQL1_Refresca(this, e);
		}

		C1.Data.C1DataRow drPagoLote;
		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (idDeposito.Text.Trim().Length == 0)
			{
				MessageBox.Show("No existen registros para editar", "Información");
				return;
			}

			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene Acceso a Editar Pagos",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.barraDato1.EditarRegistro();
			drPagoLote = cdsPagoLote.TableViews["PagoLote"].Rows.Find(int.Parse(idDeposito.Value.ToString()));
			this.Habilita(true);
			chkAsiento.Enabled = false;
		}

		private void NumeroPago()
		{
			if (this.barraDato1.bNuevo)
			{
				int iTipo = 0;
				if (IdTipoFactura != 1) iTipo = 1;
				string stNumero = Funcion.NumeraFactura(this.cdsCompraTabla, 21 + iTipo);
				if (stNumero.Trim().Length > 0) drPagoLote["Codigo"] = stNumero;

				if (iTipo != 1) return;

				if (!chkAsiento.Checked)
				{
					int iNumCheque = Contabilidad.NumeroCheque(
						cdsCuentaCorr, (int) cmbCuentaCorr.Value, true);
					if (iNumCheque != -1)	drPagoLote["Cheque"] = iNumCheque.ToString();
				}
			}
		}

		private bool Verifica()
		{
			bool bOk = true;
			this.BorraErrores();
			if (cmbHasta.Value == System.DBNull.Value)
			{
				this.errorProvider.SetError(cmbHasta, "Ingrese Fecha");
				bOk = false;
			}
			if((int) this.cmbCuentaCorr.Value == 0)
			{
				this.errorProvider.SetError(this.cmbCuentaCorr, "Ingrese Cuenta corriente");
				bOk = false;
			}
			return bOk;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			double dTotal = (double) txtTotal.Value;
			if (!Verifica()) return;

			drPagoLote = cdsPagoLote.TableViews["PagoLote"].Rows[0];
			NumeroPago();
		
//			string stNumero = "Update Pago Set idDeposito = " + idDeposito.Text.Trim();
//			int IdCuentaCorr = (int) cmbCuentaCorr.Value;
//			int IdBanco = (int) cmbBanco.Value;
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
//			{
//				int IdPago = (int) dr.Cells["idPago"].Value;
//				string stSelect = stNumero + ", idCuentaCorriente = " + IdCuentaCorr.ToString();
//				if (IdBanco > 0)
//					stSelect += ", idBanco = " + IdBanco.ToString();
//				stSelect += " Where idPago = "	+ IdPago.ToString();
//				Funcion.EjecutaSQL(cdsCliente, stSelect, true);
//			}

			if (!barraDato1.bNuevo)
			{
				drPagoLote["Editar"] = true;
				drPagoLote["FechaEditado"] = DateTime.Now;
			}
			drPagoLote["idTipoFactura"] = IdTipoFactura;
			drPagoLote["Total"] = dTotal;
			barraDato1.GrabaRegistro();
			cdsPago.Update();
			udsPagos.Rows.Clear();
			barraDatoSQL1_Refresca(this, e);

			this.Habilita(false);
			this.ultraGrid1.Enabled = true;
			this.ultraGrid2.Enabled = true;
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Pago.idDeposito} = " + idDeposito.Value.ToString();
			string stRep = "AbonoLotesDia.Rpt";
//			if (IdTipoFactura == 1) stRep = "CobroLotesDia.Rpt";
				Reporte miRep = new Reporte(stRep, stFiltro);
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		C1.Data.C1DataRow drSeteoF;
		Acceso miAcceso;
		private void PagoLotes_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'PagoLoteDia'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
			txtDeseado.Width = 0;
			cmbDesde.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
			cmbTipo.Value = 0;
			miAcceso =new Acceso("04090L5", this.cdsSeguridad.StorageDataSet.Tables["Seguridad"]);

			barraDato1.stFiltroExt = " And idTipoFactura = " 
				+ IdTipoFactura.ToString() + " And IdDeposito > 0";
			string stFiltro="";
			if (IdTipoFactura == 1) 
			{
				stFiltro = "Proveedor = 0 Or Ambos <> 0";
				this.Text = "Depósitos por Lotes";
			}
			else 
			{
				this.Text = "Cuentas por Pagar por Lotes";
				label1.Text = "Prov.";
				stFiltro = " Proveedor <> 0 Or Ambos <> 0";
			}

			C1.Data.FilterCondition fcCliente = new 
				C1.Data.FilterCondition(cdsCliente, "Cliente", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcCliente);
			cdsCliente.Fill(fcTotal, false);

			if (iIdPagoLote > 0)
			{
				barraDato1.IdRegistro = iIdPagoLote;
				barraDato1.ProximoId(5);
			}
			else
				this.barraDato1.PosUltima();
			this.barraDato1.HabilitaControles(false);
			this.idDeposito.Width = 0;
			this.idPagoLote.Width = 0;
			this.txtIdAsiento.Width = 0;
			drSeteoF = this.cdsSeteoF.TableViews["SeteoF"].Rows[0];
			this.ultraGrid1.Enabled = true;
			this.ultraGrid2.Enabled = true;
			Cursor = Cursors.Default;
		}

		private void cdsPagoLote_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			if (e.DataMember.ToString() == "_PagoLote")
			{
				this.drPagoLote = e.Row;
			}
		}


		private void GrupoCxC(int IdAsiento, double dTotal)
		{
			int IdCuenta = 0;
			double dValor = 0;
			C1.Data.C1DataRow drAsigna;
			C1.Data.C1DataRow drCC = this.cdsCuentaCorr.TableViews[
				"CuentaCorriente"].Rows.Find((int) this.cmbCuentaCorr.Value);
			if ((int) this.cmbCuentaCorr.Value != 0 || drCC["idCuenta"] != null || (int) drCC["idCuenta"] == 0)
				IdCuenta = (int) drCC["idCuenta"];
			else
			{
				drAsigna = this.cdsAsignaCuenta.TableViews["AsignaCuenta"].Rows[0];
				IdCuenta = (int) drAsigna["idCuenta"]; // Caja o Bancos en Ventas
			}

			C1.Data.C1DataRow drDetAsi;
			int iCuenta1 = 0;
			if (IdTipoFactura == 1) 
			{
				if (!chkDesglosa.Checked) 
				{
					drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
					drDetAsi["Debe"] = dTotal;
					drDetAsi["idCuenta"] = IdCuenta;
					drDetAsi["idAsiento"] = IdAsiento;
				}
				else
				{
					string stCuentaGrupo = "select sum(IsNull(pago.pago, 0)), Max(cuentacorriente.idCuenta) "
						+ "from pago inner join cuentacorriente on "
						+ "pago.idcuentacorriente = cuentacorriente.idcuentacorriente "
						+ "Where Pago.idDeposito = " 
						+ this.idDeposito.Value.ToString();
					if (drSeteoF["AgrupaReferencia"] != DBNull.Value && (bool) drSeteoF["AgrupaReferencia"] == true)
						stCuentaGrupo += " group by pago.Referencia ";
					else
						stCuentaGrupo += " group by pago.idcuentacorriente ";
				
					SqlDataReader miGrupo = Funcion.rEscalarSQL(cdsCliente, stCuentaGrupo, true);
					while (miGrupo.Read())
					{
						drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
						drDetAsi["Debe"] = miGrupo.GetDouble(0);
						drDetAsi["idCuenta"] = miGrupo.GetInt32(1);
						drDetAsi["idAsiento"] = IdAsiento;
					}
					miGrupo.Close();
				}
			}
			else iCuenta1 = IdCuenta;

			#region Detalle en asiento por cada linea
			string strSelect = "";
			if (chkDesglosa.Checked) 
			{
				string stNombre = "";
				if (chkIncluyenombre.Checked)
					stNombre += ", MIN(ISNULL(Cliente.Nombre, '')), MIN(IsNull(Compra.Numero, ''))";
				strSelect = "SELECT SUM(Pago.Pago), " +
					"MIN(IsNull(ClienteGrupo.idCuenta, 0)) " + stNombre + 
					" FROM Pago " +
					"Inner Join Compra ON Pago.idCompra = Compra.idCompra " +
					"Inner Join Cliente ON Compra.idCliente = Cliente.idCliente " +
					"Full Join ClienteGrupo ON ClienteGrupo.idGrupoCliente = " +
					"Cliente.idGrupoCliente Where Pago.idDeposito = " + 
					this.idDeposito.Value.ToString() +
					" Group by Pago.idPago ";
			}
			else
			{
				strSelect = "SELECT SUM(Pago.Pago), " +
					"MIN(ClienteGrupo.idCuenta) FROM Pago " +
					"Inner Join Compra ON Pago.idCompra = Compra.idCompra " +
					"Inner Join Cliente ON Compra.idCliente = Cliente.idCliente " +
					"Full Join ClienteGrupo ON ClienteGrupo.idGrupoCliente = " +
					"Cliente.idGrupoCliente Where Pago.idDeposito = " + 
					this.idDeposito.Value.ToString() +
					" Group by idCuenta Order by idCuenta";
			}
			SqlDataReader miReader = 
				Funcion.rEscalarSQL(cdsCliente, strSelect, true);
			try
			{
				while (miReader.Read())
				{
					dValor = miReader.GetDouble(0);
					IdCuenta = 0;
					if (miReader.GetValue(1) != System.DBNull.Value)
						IdCuenta =  miReader.GetInt32(1);
					if (IdCuenta == 0) 
					{
						int iTipo = 3;
						if (IdTipoFactura != 1) iTipo = 12;
						drAsigna = cdsAsignaCuenta.TableViews["AsignaCuenta"].Rows[iTipo];
						IdCuenta = (int) drAsigna["idCuenta"]; // Cuentas por cobrar
					}

					drDetAsi = this.cdsAsiento.TableViews["AsientoDetalle"].AddNew();
					if (IdTipoFactura == 1) drDetAsi["Haber"] = dValor;
					else drDetAsi["Debe"] = dValor;
					drDetAsi["idCuenta"] = IdCuenta;
					drDetAsi["idAsiento"] = IdAsiento;
					if (chkDesglosa.Checked)
					{
						if (chkIncluyenombre.Checked)
						{
							string stDetalle = miReader.GetString(2).Trim()
								+ " " + miReader.GetString(3).Trim();
							drDetAsi["Detalle"] = stDetalle;
						}
					}
				}
			}
			catch(System.Exception ex) 
			{
				MessageBox.Show(ex.Message, "Detalle de asiento");
			}
			miReader.Close();
		
			#endregion // Fin de Detalle de asiento por pagos no detallados

			if (IdTipoFactura != 1) 
			{
				drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
				drDetAsi["Haber"] = dTotal;
				drDetAsi["idCuenta"] = iCuenta1;
				drDetAsi["idAsiento"] = IdAsiento;
			}
			this.cdsPagoLote.Schema.Connections[0].Close();
			this.cdsAsiento.Update();
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar) 
			{
				MessageBox.Show("No puede crear un asiento mientras esta editando",
					"Información");
				return;
			}
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Crear Asientos", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			bool bRepetido = false;
			if (txtIdAsiento.Value != null && int.Parse(txtIdAsiento.Value.ToString()) > 0)
			{
				string stSelect = "SELECT COUNT(*) FROM Asiento WHERE idAsiento = "+
					this.txtIdAsiento.Value.ToString().Trim();
				int iCuenta = 0;
				iCuenta = Funcion.iEscalarSQL(cdsAsiento, stSelect);

				if (iCuenta > 0)
				{
					if (DialogResult.No == MessageBox.Show("Asiento ya existe. ¿Desea sobreescribirlo?", 
						"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk))
						return;
					// Si desea sobrrescribir
					string stBorrar = "DELETE FROM AsientoDetalle WHERE idAsiento = "
						+ this.txtIdAsiento.Value.ToString();
					Funcion.EjecutaSQL(cdsAsiento, stBorrar);
					bRepetido = true;
				}
			}

			// Crea asiento
			C1.Data.C1DataRow drAsiento;
			if (bRepetido)
			{
				string stFiltro = "idAsiento = " + txtIdAsiento.Text.ToString();
				C1.Data.FilterCondition fcPadre = new 
					C1.Data.FilterCondition(cdsAsiento, "Asiento", stFiltro);
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcPadre);

				C1.Data.FilterCondition fcHija = new 
					C1.Data.FilterCondition(cdsAsiento, "AsientoDetalle", "idAsiento = 0");
				fcTotal.Add(fcHija);
			
				this.cdsAsiento.Fill(fcTotal, false);

				drAsiento = cdsAsiento.TableViews["Asiento"].Rows[0];
			}
			else
			{
				drAsiento = cdsAsiento.TableViews["Asiento"].AddNew();
			}

			drAsiento["Fecha"] = (DateTime) cmbHasta.Value;
			drAsiento["Notas"] = txtNotas.Text;

			string stBuscaTipoA_E = "SELECT Asiento FROM CuentaCorriente "
				+ "WHERE idCuentaCorriente = "
				+ cmbCuentaCorr.Value.ToString();
			bool bAsiento = Funcion.bEscalarSQL(cdsPagoLote, stBuscaTipoA_E, false);
			if (bAsiento == true) 
			{
				drAsiento["idTipoAsiento"] = 1; // asiento
			}
			else
			{
				if (IdTipoFactura == 1)	drAsiento["idTipoAsiento"] = 2; // Ingreso
				else drAsiento["idTipoAsiento"] = 3; // Egreso
			}
			if (chkAsiento.Checked) 
				drAsiento["idTipoAsiento"] = 1;

			drAsiento["Descripcion"] = this.cmbCliente.Text;
			//			drAsiento["idProyecto"] = (int) this.cmbProyecto.EditValue;
			//			drAsiento["idSubProyecto"] = (int) this.cmbSubProyecto.EditValue;

			if ((bool) drSeteoF["CopiarNota"] == true && this.txtNotas.Value != null)
				drAsiento["Notas"] = this.txtNotas.Value;

			if ((bool) drSeteoF["CopiarNumero"] == true && this.txtCodigo.Value != null)
				drAsiento["Numero"] = this.txtCodigo.Value;

			// Crea el codigo del asiento
			if (!bRepetido)
				drAsiento["CodAsiento"] = Contabilidad.NumeroAsiento(this.cdsAsiento, (int) drAsiento["idTipoAsiento"]);
			string stFactura = "";
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
				stFactura += dr.Cells["Concepto"].Value.ToString().Trim() + " ";
			drAsiento["Notas"] += stFactura;
			if (this.cmbCliente.Text.Length == 0)
				drAsiento["Descripcion"] = "Deposito por Lotes";
			else
				drAsiento["Descripcion"] = cmbCliente.Text;

			// Actualiza el codigo del asiento en cada uno de los pagos
			double dTotal = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				dTotal += (double) dr.Cells["Pago"].Value;
			}
			dTotal = Math.Round(dTotal, 2);
			drAsiento["Total"] = dTotal;
			drAsiento["TotalH"] = dTotal;
			this.cdsAsiento.Update();

			// Creacion de Detalle de Asiento
			this.GrupoCxC((int) drAsiento["idAsiento"], dTotal);

			//			drPagoLote = cdsPagoLote.TableViews["PagoLote"].Rows.Find(int.Parse(idDeposito.Value.ToString()));
			drPagoLote = cdsPagoLote.TableViews["PagoLote"].Rows.Find(int.Parse(idPagoLote.Value.ToString()));
			drPagoLote["idAsiento"] = drAsiento["idAsiento"];
			drPagoLote["CodAsiento"] = drAsiento["CodAsiento"];
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				dr.Cells["NumeroIngreso"].Value = drAsiento["CodAsiento"];
//				dr.Cells["Notas"].Value = drAsiento["CodAsiento"];
				dr.Cells["idAsiento"].Value = drAsiento["idAsiento"];
			}
			cdsPagoLote.Update();
			cdsPago.Update();
			
			MessageBox.Show("Asiento Generado", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
		CultureInfo us = new CultureInfo("en-US");

		private void btAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (int.Parse(this.txtIdAsiento.Value.ToString()) == 0)
			{
				MessageBox.Show("No existe un asiento creado de estos Pagos");
				return;
			}
			Asiento miAsiento = new Asiento(int.Parse(this.txtIdAsiento.Value.ToString()));
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
		}

		C1.Data.C1DataRow drSeteo;
		private void btRetencion_Click(object sender, System.EventArgs e)
		{
			drSeteo = this.cdsSeteo.TableViews["Seteo"].Rows[0];
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la factura para crear una retención",
					"Información");
				return;
			}
			Cursor = Cursors.WaitCursor;
			int IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;

			C1.Data.C1DataRow drCompra = cdsCompra.TableViews["Compra"].Rows.Find(IdCompra);
			if (this.IdTipoFactura == 4) // Compra
			{
				if (drSeteo["FormMarzo"] == System.DBNull.Value 
					|| (bool) drSeteo["FormMarzo"] == false)
				{
//					using (Egreso miEgreso = new Egreso())
//					{
//						miEgreso.IdFactura = IdCompra;
//						miEgreso.drFactura = drCompra;
//						miEgreso.ShowDialog();
//					}
				}
				else
				{
//					using(EgresoAT miEgreso = new EgresoAT())
//					{
//						miEgreso.IdFactura = IdCompra;
//						miEgreso.drFactura = drCompra;
//						miEgreso.txtIdPagoLote.EditValue = int.Parse(idDeposito.Value.ToString());
//						miEgreso.ShowDialog();
//					}
				}
			}

			if (this.IdTipoFactura == 1) // Venta
			{
				if (drSeteo["FormMarzo"] == System.DBNull.Value 
					|| (bool) drSeteo["FormMarzo"] == false)
				{
					using(Ingreso miIngreso = new Ingreso())
					{
						miIngreso.IdFactura = IdCompra;
						miIngreso.drFactura = drCompra;
						miIngreso.ShowDialog();
					}
				}
				else
				{
					using(IngresoAT miIngreso = new IngresoAT())
					{
						miIngreso.IdFactura = IdCompra;
						miIngreso.drFactura = drCompra;
						miIngreso.txtIdPagoLote.EditValue = int.Parse(idDeposito.Value.ToString());
						miIngreso.ShowDialog();
					}
				}
			}
			btFiltro_Click(this, e);
			Cursor = Cursors.Default;
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "idCliente = " + this.cmbCliente.Value.ToString()
				+ " And Saldo >= 0.01 And Borrar = 0";
			if (IdTipoFactura == 1)
				stFiltro += " And (idTipoFactura = 1 OR idTipoFactura = 24 Or "
					+ "idTipoFactura = 11 Or idTipoFactura = 27)";
			else
				stFiltro += " And (idTipoFactura = 4 OR idTipoFactura = 25 Or "
          + "idTipoFactura = 26 Or idTipoFactura = 28)";

			DateTime dtFechaIni = (DateTime) cmbDesde.Value;
			DateTime dtFechaFin = (DateTime) cmbHasta.Value;
			stFiltro += " And FechaVencimiento >= '" + dtFechaIni.ToString("yyyyMMdd") 
				+ "' And FechaVencimiento < '" + dtFechaFin.AddDays(1).ToString("yyyyMMdd") + "'";
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsCompra, "Compra", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

				C1.Data.FilterCondition fcHija = new 
					C1.Data.FilterCondition(this.cdsCompra, "DetCompra", "idCompra = 0");
				fcTotal.Add(fcHija);
			
			this.cdsCompra.Fill(fcTotal, false);
			Cursor = Cursors.Default;
//			MessageBox.Show("Finalizado", "Información");
		}

	
		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

			string stbProveedor = "";
			switch(IdTipoFactura)
			{
				case 1:
				case 3:
				case 5:
				case 13:
				case 15:
					stbProveedor = "False";
					break;
				case 2:
				case 4:
				case 6:
				case 14:
				case 16:
					stbProveedor = "True";
					break;
			}
			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsClientes"];

			if (stbProveedor == "True") stbProveedor = " <> 0";
			else stbProveedor = " = 0";

			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Cliente"],
				"[Ambos] <> 0 OR [Proveedor] " + stbProveedor));
		}

		private void cdsPagoLote_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPagoLote.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsFormaPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsFormaPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCuentaCorr_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuentaCorr.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsAsiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsiento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCompraTabla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompraTabla.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsAsignaCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAsignaCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			int IdDeposito = -1;
			if (idDeposito.Text.Length > 0)
				IdDeposito = int.Parse(idDeposito.Text);
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsPago, "Pago", "idDeposito = " + IdDeposito.ToString());
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			cdsPago.Fill(fcTotal, false);
		}

		private void btAgregarTodo_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void btRetencion_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la factura para crear el Asiento de la retencion",
					"Información");
				return;
			}
			// Crea asiento de Retencion
			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Crear Asientos", "Seguridad", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea crear Asiento de Retención?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;

			Cursor = Cursors.WaitCursor;

			int IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
			Contabilidad.AsientoRetencion(cdsCompra, IdCompra, 1);
			Cursor = Cursors.Default;
//				int IdProyecto = 0;
//				int IdSubProyecto = 0;
//				string stNumero = "";
//				if (ultraGrid1.ActiveRow.Cells["idProyecto"].Value != System.DBNull.Value)
//					IdProyecto = (int) ultraGrid1.ActiveRow.Cells["idProyecto"].Value;
//				if (ultraGrid1.ActiveRow.Cells["idSubProyecto"].Value != System.DBNull.Value)
//					IdSubProyecto = (int) ultraGrid1.ActiveRow.Cells["idSubProyecto"].Value;
//				if ((bool) drSeteoF["CopiarNumero"] == true)
//					if (ultraGrid1.ActiveRow.Cells["Numero"].Value != System.DBNull.Value)
//						stNumero = (string) ultraGrid1.ActiveRow.Cells["Numero"].Value;
//
//				string stCodAsiento = Contabilidad.NumeroAsiento(cdsAsiento, 1);
//				string stSelect = "Insert Into Asiento "
//					+ "(Descripcion, Fecha, idTipoAsiento, CodAsiento, Numero"
//					+ ", idProyecto, idSubProyecto) VALUES ("
//					+ "'Retencion', '" +  DateTime.Today.ToString("yyyMMdd")+ "', 1, '"
//					+ stCodAsiento + "', '" + stNumero + "', " + IdProyecto.ToString()
//					+ ", " + IdSubProyecto.ToString() + ")";
//
//				Funcion.EjecutaSQL(cdsAsiento, stSelect);
//				stSelect = "Select IdAsiento From Asiento Where CodAsiento = '" 
//					+ stCodAsiento + "'";
//				int IdPrinc = Funcion.iEscalarSQL(cdsAsiento, stSelect);
//				// Crea el codigo del asiento
//				Funcion.RetencionCuenta(cdsAsiento, IdCompra, 1, IdPrinc);
//				MessageBox.Show("Asiento de Retencion Generado", "Información");
//			}
//			else
//			{
//				MessageBox.Show("No existe retencion en la factura para generar un asiento");
//				return;
//			}

		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Pago que desea ver el detalle", "Información");
				return;
			}
			int IdPago = (int) ultraGrid1.ActiveRow.Cells["idPago"].Value;
//			Pago miPago = new Pago(IdPago, false, 0);
//			miPago.MdiParent = MdiParent;
//			miPago.Show();
		}

		private void ultraGrid2_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.IsInEditMode)
				if (e.Cell.Column.ToString() == "Pago")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						if (drSeteoF["PagoNegativo"] == System.DBNull.Value ||
							(bool) drSeteoF["PagoNegativo"] == false)
						{
							string stSelect = "Select Saldo From Compra Where idCompra = " 
								+ e.Cell.Row.Cells["idCompra"].Value.ToString();
							double dSaldo = Funcion.dEscalarSQL(cdsCompra, stSelect, true);
							dSaldo = Math.Round(dSaldo, 2);
							if ((double) e.Cell.Value > dSaldo)
							{
								MessageBox.Show("Pago Excede Saldo", "Información");
								e.Cell.Row.Cells["Pago"].Value = dSaldo;
							}
              txtTotal_DoubleClick(this, e);
						}  // Fin de Pago Negativo = falso
					}

		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void ultraGrid2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid2.ActiveCell == null) return;
			try
			{
				if (ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid2.ActiveCell.DroppedDown == false)
							ultraGrid2.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid2.ActiveCell.DroppedDown) return;

						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid2.ActiveCell.DroppedDown) return;
			
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid2.ActiveCell.DroppedDown == false)
							ultraGrid2.ActiveCell.DroppedDown = true;
						//						ultraGrid2.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
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
						break;
					case (int) Keys.Insert:
						if (e.Control)
						{
							System.Windows.Forms.Clipboard.SetDataObject(ultraGrid1.Selected.Cells);
							MessageBox.Show("Filas Copiadas");
						}

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

		private void txtTotal_DoubleClick(object sender, System.EventArgs e)
		{
			double dPago = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if (dr.Cells["Pago"].Value != System.DBNull.Value)
					dPago += (double) dr.Cells["Pago"].Value;
			}
			txtTotal.Value = dPago;	
		}

		private void btPago_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			if (cmbHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese Fecha Ultima");
				return;
			}
			string stSelect = "select IsNull(Pago.idFormaPago, 0), "
				+ "IsNull(Pago.Fecha, '19010101'), IsNull(Pago.Pago, 0), "
				+ "IsNull(Pago.Vencimiento, '19010101'), "
				+ "Isnull(Cliente.Nombre, ''), IsNull(Pago.idBanco, 0), "
				+ "IsNull(Compra.Total, 0), IsNull(Pago.Numero, ''), "
				+ "IsNull(Pago.idPago, 0), IsNull(Pago.idCuentaCorriente, 0), "
				+ "IsNull(Pago.Pendiente, 0), IsNull(Pago.Referencia, ''), IsNull(Pago.Cheque, 0), "
				+ "IsNull(Compra.Numero, 0), IsNull(Pago.Cuenta, 0) "
				+ "from pago inner join compra on pago.idcompra = compra.idcompra "
				+ "Inner Join Cliente On Cliente.idCliente = Compra.idCliente "
				+ "where ";
			DateTime dtFechaIni = (DateTime) cmbDesde.Value;
			DateTime dtFechaFin = (DateTime) cmbHasta.Value;
			string stFecha = "Pago.Fecha";
			if (chkVencimiento.Checked) stFecha = "Pago.Vencimiento";

			stSelect += stFecha + " >= '" + dtFechaIni.ToString("yyyyMMdd") 
				+ "' And " + stFecha + " < '" + dtFechaFin.AddDays(1).ToString("yyyyMMdd") 
				+ "' And IsNull(IdDeposito, 0) = 0 ";
			//				+ "Order by Cliente.Nombre ";
//			if ((int) cmbTipo.Value == 1)
//				stSelect += " And Pago.Pendiente = 1 ";
//			if ((int) cmbTipo.Value == 0)
//				stSelect += " And Pago.Pendiente = 0 ";
			if (!chkCliente.Checked)
				stSelect += " And Compra.idCliente = " + this.cmbCliente.Value.ToString();
			if (IdTipoFactura == 1)
				stSelect += " And (idTipoFactura = 1 OR idTipoFactura = 24 Or "
					+ "idTipoFactura = 11 Or idTipoFactura = 27)";
			else
				stSelect += " And (idTipoFactura = 4 OR idTipoFactura = 25 Or "
					+ "idTipoFactura = 26 Or idTipoFactura = 28)";
			if (!chkCuentaCorr.Checked)
				stSelect += " And Pago.idCuentaCorriente = " + cmbCuentaCorr.Value.ToString();
			if (!chkFormaPago.Checked)
				stSelect += " And Pago.idFormaPago = " + cmbFormaPago.Value.ToString();
			if (!chkTdoBanco.Checked)
				stSelect += " And Pago.idBanco = " + cmbBanco.Value.ToString();

			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCliente, stSelect, true);
			udsPagos.Rows.Clear();
			int iNumFilas = -1;
			while (miReader.Read())
			{
				int IdFormaPago = miReader.GetInt32(0);
				dtFechaIni = miReader.GetDateTime(1);
				double dPago = miReader.GetDouble(2);
				dtFechaFin = miReader.GetDateTime(3);
				string stNombre = miReader.GetString(4);
				int IdBanco = miReader.GetInt32(5);
				double dTotal = miReader.GetDouble(6);
				string stNumero = miReader.GetString(7);
				int IdPago = miReader.GetInt32(8);
				int IdCuentaCorr = miReader.GetInt32(9);
				bool bPendiente = miReader.GetBoolean(10);
				string stReferencia = miReader.GetString(11);
				string stCheque = miReader.GetString(12);
				string stFactura = miReader.GetString(13);
				string stCuenta = miReader.GetString(13);

				iNumFilas ++;
				udsPagos.Rows.Add();
				udsPagos.Rows[iNumFilas]["idFormaPago"]=IdFormaPago;
				udsPagos.Rows[iNumFilas]["Fecha"]=dtFechaIni;
				udsPagos.Rows[iNumFilas]["Pago"]=dPago;
				udsPagos.Rows[iNumFilas]["Vecimiento"]=dtFechaFin;
				udsPagos.Rows[iNumFilas]["Nombre"]=stNombre;
				udsPagos.Rows[iNumFilas]["Numero"]=stNumero;
				udsPagos.Rows[iNumFilas]["idBanco"]=IdBanco;
				udsPagos.Rows[iNumFilas]["Total"]=dTotal;
				udsPagos.Rows[iNumFilas]["idPago"]=IdPago;
				udsPagos.Rows[iNumFilas]["idCuentaCorriente"]=IdCuentaCorr;
				udsPagos.Rows[iNumFilas]["Pendiente"]=bPendiente;
				udsPagos.Rows[iNumFilas]["Referencia"]=stReferencia;
				udsPagos.Rows[iNumFilas]["Cheque"]=stCheque;
				udsPagos.Rows[iNumFilas]["Factura"]=stFactura;
				udsPagos.Rows[iNumFilas]["Notas"]=stCuenta;
			}
			miReader.Close();
			Cursor = Cursors.Default;
		}

		private void cmbCuentaCorr_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (DialogResult.No == MessageBox.Show("Desea Copiar Cuenta Corriente a Pagos Seleccionados",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
			{
				if (dr.Cells["idPago"].Value != System.DBNull.Value)
				{
          dr.Cells["idCuentaCorriente"].Value = (int) cmbCuentaCorr.Value;
				}
			}
		}

		private void btAgregar_Click(object sender, System.EventArgs e)
		{
			if (!VerificaAgregar()) return;
			errorProvider.SetError(ultraGrid1, "");
			if (ultraGrid1.Selected.Rows.Count == 0)
				ultraGrid1.ActiveRow.Selected = true;

			string stNumero = "Update Pago Set idDeposito = " + idDeposito.Text.Trim();
			int IdCuentaCorr = (int) cmbCuentaCorr.Value;
			int IdBanco = 0;
			if (!chkTdoBanco.Checked) IdBanco = (int) cmbBanco.Value;
			Infragistics.Win.UltraWinGrid.UltraGridRow[] drElimina = new UltraGridRow[500];
			int i = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
			{
				int IdPago = (int) dr.Cells["idPago"].Value;
				string stSelect = stNumero + ", idCuentaCorriente = " + IdCuentaCorr.ToString();
				if (IdBanco > 0)
					stSelect += ", idBanco = " + IdBanco.ToString();
				stSelect += " Where idPago = "	+ IdPago.ToString();
				Funcion.EjecutaSQL(cdsCliente, stSelect, true);
				drElimina[i++] = dr;
			}

			for(int j=0; j<i; j++) drElimina[j].Delete(false);

			barraDatoSQL1_Refresca(this, e);
			ultraGrid2.Refresh();
			txtTotal_DoubleClick(this, e);
		}

		private void PagoLoteDia_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'PagoLoteDia'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}

	}
}
