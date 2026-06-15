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
	/// Descripción breve de AgroLabor.
	/// </summary>
	public class AgroLabor : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsAgroLabor;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataSet cdsPersonal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataView cdvPersonal;
		private C1.Data.C1DataSet cdsUnidad;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbDesde;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private DevExpress.XtraEditors.SimpleButton btReporte;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArt;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkArticulo;
		private DevExpress.XtraEditors.SimpleButton btIngreso;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private DevExpress.XtraEditors.SimpleButton btGenera;
		private C1.Data.C1DataSet cdsRolConcepto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAportable;
		private DevExpress.XtraEditors.SimpleButton btExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public AgroLabor()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PaqueteUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CambiarPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredefCR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoFecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aportable");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			this.cdsAgroLabor = new C1.Data.C1DataSet();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbDesde = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsUnidad = new C1.Data.C1DataSet();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.btReporte = new DevExpress.XtraEditors.SimpleButton();
			this.chkPersonal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbArt = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkArticulo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btIngreso = new DevExpress.XtraEditors.SimpleButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsRolConcepto = new C1.Data.C1DataSet();
			this.btGenera = new DevExpress.XtraEditors.SimpleButton();
			this.chkAportable = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btExcel = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			((System.ComponentModel.ISupportInitialize)(this.cdsAgroLabor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUnidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsAgroLabor
			// 
			this.cdsAgroLabor.BindingContextCtrl = this;
			this.cdsAgroLabor.DataLibrary = "LibPersonal";
			this.cdsAgroLabor.DataLibraryUrl = "";
			this.cdsAgroLabor.DataSetDef = "dsAgro_Rol";
			this.cdsAgroLabor.EnforceConstraints = false;
			this.cdsAgroLabor.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAgroLabor.Name = "cdsAgroLabor";
			this.cdsAgroLabor.SchemaClassName = "LibPersonal.DataClass";
			this.cdsAgroLabor.SchemaDef = null;
			this.cdsAgroLabor.BeforeFill += new C1.Data.FillEventHandler(this.cdsAgroLabor_BeforeFill);
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
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
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulo";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.RowFilter = "idClase=2";
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// ultraTextEditor3
			// 
			this.ultraTextEditor3.Location = new System.Drawing.Point(440, 8);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(72, 21);
			this.ultraTextEditor3.TabIndex = 83;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(384, 8);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(43, 14);
			this.ultraLabel2.TabIndex = 85;
			this.ultraLabel2.Text = "Codigo:";
			// 
			// cmbDesde
			// 
			this.cmbDesde.DateTime = new System.DateTime(2008, 12, 1, 0, 0, 0, 0);
			this.cmbDesde.Location = new System.Drawing.Point(104, 40);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.Size = new System.Drawing.Size(88, 21);
			this.cmbDesde.TabIndex = 86;
			this.cmbDesde.Value = new System.DateTime(2008, 12, 1, 0, 0, 0, 0);
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(24, 40);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(75, 14);
			this.ultraLabel3.TabIndex = 87;
			this.ultraLabel3.Text = "Fecha Desde:";
			// 
			// cdsUnidad
			// 
			this.cdsUnidad.BindingContextCtrl = this;
			this.cdsUnidad.DataLibrary = "LibFacturacion";
			this.cdsUnidad.DataLibraryUrl = "";
			this.cdsUnidad.DataSetDef = "dsArticuloTabla";
			this.cdsUnidad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsUnidad.Name = "cdsUnidad";
			this.cdsUnidad.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsUnidad.SchemaDef = null;
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(200, 40);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(71, 14);
			this.ultraLabel4.TabIndex = 90;
			this.ultraLabel4.Text = "Fecha Hasta:";
			this.ultraLabel4.Click += new System.EventHandler(this.ultraLabel4_Click);
			// 
			// cmbHasta
			// 
			this.cmbHasta.DateTime = new System.DateTime(2008, 12, 1, 0, 0, 0, 0);
			this.cmbHasta.Location = new System.Drawing.Point(280, 40);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.Size = new System.Drawing.Size(88, 21);
			this.cmbHasta.TabIndex = 89;
			this.cmbHasta.Value = new System.DateTime(2008, 12, 1, 0, 0, 0, 0);
			// 
			// btReporte
			// 
			this.btReporte.Location = new System.Drawing.Point(464, 72);
			this.btReporte.Name = "btReporte";
			this.btReporte.Size = new System.Drawing.Size(72, 24);
			this.btReporte.TabIndex = 91;
			this.btReporte.Text = "Borrar";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// chkPersonal
			// 
			this.chkPersonal.Checked = true;
			this.chkPersonal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPersonal.Location = new System.Drawing.Point(24, 72);
			this.chkPersonal.Name = "chkPersonal";
			this.chkPersonal.Size = new System.Drawing.Size(96, 20);
			this.chkPersonal.TabIndex = 92;
			this.chkPersonal.Text = "Todo personal";
			// 
			// cmbPersonal
			// 
			this.cmbPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataSource = this.cdvPersonal;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
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
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 275;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 35;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 36;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 37;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 38;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 39;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 41;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 44;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 45;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 46;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 47;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 48;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 49;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 50;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 51;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 52;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 53;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 54;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 55;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 56;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 57;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 58;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 59;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 60;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 61;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 62;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 63;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 64;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 65;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 66;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 67;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 68;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 69;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 70;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 71;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 72;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 73;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 74;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 75;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 76;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 77;
			ultraGridColumn78.Hidden = true;
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
																										 ultraGridColumn78});
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonal.Location = new System.Drawing.Point(152, 72);
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(216, 21);
			this.cmbPersonal.TabIndex = 93;
			this.cmbPersonal.ValueMember = "idPersonal";
			// 
			// cmbArt
			// 
			this.cmbArt.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArt.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArt.DataSource = this.cdvArticulo;
			ultraGridBand2.AddButtonCaption = "Personal";
			ultraGridColumn79.Header.VisiblePosition = 3;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 1;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 4;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 5;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 10;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 8;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 11;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 12;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 14;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 16;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 9;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 18;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 19;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 20;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 21;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 22;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 23;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 24;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 25;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 26;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 27;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 28;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 29;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 30;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 31;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 32;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 33;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 34;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 35;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 36;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 37;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 38;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 39;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 40;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 41;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 42;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 43;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 44;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 45;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 46;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 47;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 48;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 49;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 50;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 51;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 52;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 53;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 54;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.VisiblePosition = 55;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn128.Header.VisiblePosition = 56;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.VisiblePosition = 57;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn130.Header.VisiblePosition = 58;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn131.Header.VisiblePosition = 59;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn132.Header.VisiblePosition = 60;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn133.Header.VisiblePosition = 61;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn134.Header.VisiblePosition = 62;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn135.Header.VisiblePosition = 63;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn136.Header.VisiblePosition = 64;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 65;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn138.Header.VisiblePosition = 66;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 67;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn140.Header.VisiblePosition = 68;
			ultraGridColumn140.Hidden = true;
			ultraGridColumn141.Header.VisiblePosition = 69;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 70;
			ultraGridColumn142.Hidden = true;
			ultraGridColumn143.Header.VisiblePosition = 71;
			ultraGridColumn143.Hidden = true;
			ultraGridColumn144.Header.VisiblePosition = 72;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn145.Header.VisiblePosition = 73;
			ultraGridColumn145.Hidden = true;
			ultraGridColumn146.Header.VisiblePosition = 74;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn147.Header.VisiblePosition = 75;
			ultraGridColumn147.Hidden = true;
			ultraGridColumn148.Header.VisiblePosition = 76;
			ultraGridColumn148.Hidden = true;
			ultraGridColumn149.Header.VisiblePosition = 77;
			ultraGridColumn149.Hidden = true;
			ultraGridColumn150.Header.VisiblePosition = 78;
			ultraGridColumn150.Hidden = true;
			ultraGridColumn151.Header.VisiblePosition = 79;
			ultraGridColumn151.Hidden = true;
			ultraGridColumn152.Header.VisiblePosition = 80;
			ultraGridColumn152.Hidden = true;
			ultraGridColumn153.Header.VisiblePosition = 81;
			ultraGridColumn153.Hidden = true;
			ultraGridColumn154.Header.VisiblePosition = 82;
			ultraGridColumn154.Hidden = true;
			ultraGridColumn155.Header.VisiblePosition = 83;
			ultraGridColumn155.Hidden = true;
			ultraGridColumn156.Header.VisiblePosition = 84;
			ultraGridColumn156.Hidden = true;
			ultraGridColumn157.Header.VisiblePosition = 85;
			ultraGridColumn157.Hidden = true;
			ultraGridColumn158.Header.VisiblePosition = 86;
			ultraGridColumn158.Hidden = true;
			ultraGridColumn159.Header.VisiblePosition = 87;
			ultraGridColumn159.Hidden = true;
			ultraGridColumn160.Header.VisiblePosition = 88;
			ultraGridColumn160.Hidden = true;
			ultraGridColumn161.Header.VisiblePosition = 2;
			ultraGridColumn161.Hidden = true;
			ultraGridColumn162.Header.VisiblePosition = 89;
			ultraGridColumn162.Hidden = true;
			ultraGridColumn163.Header.VisiblePosition = 90;
			ultraGridColumn163.Hidden = true;
			ultraGridColumn164.Header.VisiblePosition = 91;
			ultraGridColumn164.Hidden = true;
			ultraGridColumn165.Header.VisiblePosition = 92;
			ultraGridColumn165.Hidden = true;
			ultraGridColumn166.Header.VisiblePosition = 93;
			ultraGridColumn166.Hidden = true;
			ultraGridColumn167.Header.VisiblePosition = 94;
			ultraGridColumn167.Hidden = true;
			ultraGridColumn168.Header.VisiblePosition = 95;
			ultraGridColumn168.Hidden = true;
			ultraGridColumn169.Header.VisiblePosition = 96;
			ultraGridColumn169.Hidden = true;
			ultraGridColumn170.Header.VisiblePosition = 97;
			ultraGridColumn170.Hidden = true;
			ultraGridColumn171.Header.VisiblePosition = 98;
			ultraGridColumn171.Hidden = true;
			ultraGridColumn172.Header.VisiblePosition = 99;
			ultraGridColumn172.Hidden = true;
			ultraGridColumn173.Header.VisiblePosition = 100;
			ultraGridColumn173.Hidden = true;
			ultraGridColumn174.Header.VisiblePosition = 101;
			ultraGridColumn174.Hidden = true;
			ultraGridColumn175.Header.VisiblePosition = 102;
			ultraGridColumn175.Hidden = true;
			ultraGridColumn176.Header.VisiblePosition = 103;
			ultraGridColumn176.Hidden = true;
			ultraGridColumn177.Header.VisiblePosition = 6;
			ultraGridColumn177.Hidden = true;
			ultraGridColumn178.Header.VisiblePosition = 7;
			ultraGridColumn178.Hidden = true;
			ultraGridColumn179.Header.VisiblePosition = 104;
			ultraGridColumn179.Hidden = true;
			ultraGridColumn180.Header.VisiblePosition = 105;
			ultraGridColumn180.Hidden = true;
			ultraGridColumn181.Header.VisiblePosition = 0;
			ultraGridColumn181.Width = 304;
			ultraGridColumn182.Header.VisiblePosition = 13;
			ultraGridColumn183.Header.VisiblePosition = 106;
			ultraGridColumn183.Hidden = true;
			ultraGridColumn184.Header.VisiblePosition = 107;
			ultraGridColumn184.Hidden = true;
			ultraGridColumn185.Header.VisiblePosition = 15;
			ultraGridColumn185.Hidden = true;
			ultraGridColumn186.Header.VisiblePosition = 17;
			ultraGridColumn186.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn186});
			this.cmbArt.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbArt.DisplayMember = "Articulo";
			this.cmbArt.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArt.Location = new System.Drawing.Point(152, 8);
			this.cmbArt.Name = "cmbArt";
			this.cmbArt.Size = new System.Drawing.Size(216, 21);
			this.cmbArt.TabIndex = 95;
			this.cmbArt.ValueMember = "idArticulo";
			// 
			// chkArticulo
			// 
			this.chkArticulo.Checked = true;
			this.chkArticulo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkArticulo.Location = new System.Drawing.Point(24, 8);
			this.chkArticulo.Name = "chkArticulo";
			this.chkArticulo.Size = new System.Drawing.Size(96, 20);
			this.chkArticulo.TabIndex = 94;
			this.chkArticulo.Text = "Todo Articulo";
			// 
			// btIngreso
			// 
			this.btIngreso.Location = new System.Drawing.Point(544, 72);
			this.btIngreso.Name = "btIngreso";
			this.btIngreso.Size = new System.Drawing.Size(72, 24);
			this.btIngreso.TabIndex = 96;
			this.btIngreso.Text = "Envia Rol";
			this.btIngreso.Click += new System.EventHandler(this.btIngreso_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(520, 8);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(26, 14);
			this.ultraLabel1.TabIndex = 98;
			this.ultraLabel1.Text = "Tipo";
			this.ultraLabel1.Click += new System.EventHandler(this.ultraLabel1_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
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
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.ultraGrid1.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.AppearancesOnly;
			this.ultraGrid1.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.AllRowsInBand;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			appearance6.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultraGrid1.DisplayLayout.Override.SummaryValueAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 112);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(690, 184);
			this.ultraGrid1.TabIndex = 99;
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.Location = new System.Drawing.Point(520, 40);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(47, 14);
			this.ultraLabel5.TabIndex = 101;
			this.ultraLabel5.Text = "Posicion";
			// 
			// ultraNumericEditor1
			// 
			this.ultraNumericEditor1.Location = new System.Drawing.Point(592, 40);
			this.ultraNumericEditor1.MaskInput = "n";
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(104, 21);
			this.ultraNumericEditor1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.ultraNumericEditor1.TabIndex = 100;
			this.ultraNumericEditor1.Value = 1;
			// 
			// cmbTipo
			// 
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataMember = "RolTipo";
			this.cmbTipo.DataSource = this.cdsRolConcepto;
			ultraGridColumn187.Header.VisiblePosition = 1;
			ultraGridColumn187.Hidden = true;
			ultraGridColumn188.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn187,
																										 ultraGridColumn188});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(592, 8);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(104, 21);
			this.cmbTipo.TabIndex = 102;
			this.cmbTipo.ValueMember = "idRolTipo";
			this.cmbTipo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipo_InitializeLayout);
			// 
			// cdsRolConcepto
			// 
			this.cdsRolConcepto.BindingContextCtrl = this;
			this.cdsRolConcepto.DataLibrary = "LibPersonal";
			this.cdsRolConcepto.DataLibraryUrl = "";
			this.cdsRolConcepto.DataSetDef = "dsRolGrupo";
			this.cdsRolConcepto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolConcepto.Name = "cdsRolConcepto";
			this.cdsRolConcepto.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolConcepto.SchemaDef = null;
			// 
			// btGenera
			// 
			this.btGenera.Location = new System.Drawing.Point(384, 72);
			this.btGenera.Name = "btGenera";
			this.btGenera.Size = new System.Drawing.Size(72, 24);
			this.btGenera.TabIndex = 103;
			this.btGenera.Text = "Mostrar";
			this.btGenera.Click += new System.EventHandler(this.btGenera_Click);
			// 
			// chkAportable
			// 
			this.chkAportable.Checked = true;
			this.chkAportable.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkAportable.Location = new System.Drawing.Point(392, 40);
			this.chkAportable.Name = "chkAportable";
			this.chkAportable.Size = new System.Drawing.Size(96, 20);
			this.chkAportable.TabIndex = 104;
			this.chkAportable.Text = "No Aportable";
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(624, 72);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(72, 24);
			this.btExcel.TabIndex = 105;
			this.btExcel.Text = "Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// AgroLabor
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(712, 309);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.chkAportable);
			this.Controls.Add(this.btGenera);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.btIngreso);
			this.Controls.Add(this.cmbArt);
			this.Controls.Add(this.chkArticulo);
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.chkPersonal);
			this.Controls.Add(this.btReporte);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraTextEditor3);
			this.Name = "AgroLabor";
			this.Text = "Agro Labor";
			this.Load += new System.EventHandler(this.AgroLabor_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsAgroLabor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsUnidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}
		DataSet miDataSet = new DataSet();
		DataView dv = new DataView();
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		DataSet miDataSetT = new DataSet();
		DataView dvT = new DataView();
		private void cdsAgroLabor_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsAgroLabor.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void AgroLabor_Load(object sender, System.EventArgs e)
		{
			cmbTipo.Value = 1;
//			barraDato1.ProximoId(4);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
//			if (!miAcceso.Eliminar)
//			{
//				MessageBox.Show("No tiene acceso a Eliminar " + stProvCliente,
//					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
//				return;
//			}

			if (MessageBox.Show("¿Desea Borrar este Registro?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) 
				return;

//			barraDato1.BorraRegistro();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
//			barraDato1.CrearRegistro();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
//			barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
//			barraDato1.EditarRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
//			barraDato1.GrabaRegistro();
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			int Tipo = (int) this.cmbTipo.Value; 
			int Aportable  =  0;
			int	IdPersonal = 0;
			int	IdArticulo = 0;
			if (!chkPersonal.Checked) 
				IdPersonal = (int) this.cmbPersonal.Value;
			if (!chkArticulo.Checked) 
				IdArticulo = (int) this.cmbArt.Value;
			if (chkAportable.Checked) 
				Aportable = 1;
			int Posicion = (int) this.ultraNumericEditor1.Value;
			if (DialogResult.No == MessageBox.Show("Desea Borrar Registro de la Tabla"  
				,"Confirmación",MessageBoxButtons.YesNo,MessageBoxIcon.Question)) return;

			string stExec = "Exec AgroRolBorrar '"
				+ dtDesde.ToString("yyyyMMdd") + "', '"
				+ dtHasta.ToString("yyyyMMdd") + "', "
				+ Tipo + ", "
				+ Posicion + ", " 
				+ IdPersonal + ", "
				+ IdArticulo + ", " 
				+ Aportable;

			Funcion.EjecutaSQL(cdsAgroLabor, stExec, true);
			Cursor = Cursors.Default;
			MessageBox.Show("Apuntamientos Borrados");


		}

		private bool Verifica()
		{
			bool bOk = true;
			errorProvider.SetError(cmbArt, "");
			errorProvider.SetError(cmbPersonal, "");

			if (!chkArticulo.Checked && cmbArt.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbArt, "Ingrese Articulo");
				bOk = false;
			}
			if (!chkPersonal.Checked && cmbPersonal.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbPersonal, "Ingrese Personal");
				bOk = false;
			}

			return bOk;
		}

		private void btIngreso_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			int Tipo = (int) this.cmbTipo.Value; 
			int Aportable  =  0;
			int	IdPersonal = 0;
			int	IdArticulo = 0;
			if (!chkPersonal.Checked) 
				IdPersonal = (int) this.cmbPersonal.Value;
			if (!chkArticulo.Checked) 
				IdArticulo = (int) this.cmbArt.Value;
			if (chkAportable.Checked) 
				Aportable = 1;
			int Posicion = (int) this.ultraNumericEditor1.Value;
			if (DialogResult.No == MessageBox.Show("Desea Enviar Registro al Rol"  
				,"Confirmación",MessageBoxButtons.YesNo,MessageBoxIcon.Question)) return;

			string stExec = "Exec AgroRolIngreso '"
				+ dtDesde.ToString("yyyyMMdd") + "', '"
				+ dtHasta.ToString("yyyyMMdd") + "', "
				+ Tipo + ", "
				+ Posicion + ", " 
				+ IdPersonal + ", "
				+ IdArticulo + ", " 
				+ Aportable;

			Funcion.EjecutaSQL(cdsAgroLabor, stExec, true);
			MessageBox.Show("Labores generadas al rol", "Información",
			MessageBoxButtons.OK, MessageBoxIcon.Information); 

			Cursor = Cursors.Default;
		}

		private void ultraLabel1_Click(object sender, System.EventArgs e)
		{
		
		}			

		private void btGenera_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaIni = (DateTime) this.cmbDesde.Value;
			DateTime dtFechaFin = (DateTime) this.cmbHasta.Value;
			int	IdPersonal = 0;
			int	IdArticulo = 0;
			int NoAportable = 0;
			if (!chkPersonal.Checked) 
					IdPersonal = (int) this.cmbPersonal.Value;
			if (!chkArticulo.Checked) 
					IdArticulo = (int) this.cmbArt.Value;
			if (chkAportable.Checked) 
					NoAportable = 1;
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirRoles);	
			string stProcedimiento = "exec AgroLaborGenera" 
				+ " '" + dtFechaIni.ToString("yyyyMMdd")
				+ "', '" + dtFechaFin.ToString("yyyyMMdd")
				+ "', " + IdPersonal + ", " + IdArticulo
				+ ", " + NoAportable;
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);

			//			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSetG.Reset();
			miDataSetG.Clear();
			try
			{
				myData.Fill(miDataSetG);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}

			dvG = miDataSetG.Tables[0].DefaultView;
			ultraGrid1.DataSource = dvG;

			#region Visibilidad
			ultraGrid1.DisplayLayout.Bands[0].Columns["idPersonal"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idUnidad"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idProyecto"].Hidden = true;

			#endregion Visibilidad

		}

		private void cmbTipo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid1, "AgroRol.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "AgroRol.xls");

		}

		private void ultraLabel4_Click(object sender, System.EventArgs e)
		{
		
		}


	}
}
