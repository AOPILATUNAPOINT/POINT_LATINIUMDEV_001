using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid; 
using C1.Data;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de GuiaRemisionG.
	/// </summary>
	public class GuiaRemisionG : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsGuiaRemisionG;
		private C1.Data.C1DataView cdvGuiaRemisionG;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private DevExpress.XtraEditors.CheckEdit chkVenta;
		private DevExpress.XtraEditors.CheckEdit checkCompra;
		private DevExpress.XtraEditors.CheckEdit checkTransformacion;
		private DevExpress.XtraEditors.CheckEdit checkConsignacion;
		private DevExpress.XtraEditors.CheckEdit checkTraslado;
		private DevExpress.XtraEditors.CheckEdit checkDevolucion;
		private DevExpress.XtraEditors.CheckEdit checkImportacion;
		private DevExpress.XtraEditors.CheckEdit checkExportacion;
		private DevExpress.XtraEditors.CheckEdit checkOtros;
		private DevExpress.XtraEditors.DateEdit dtFinicioT;
		private DevExpress.XtraEditors.DateEdit dtFTermT;
		private DevExpress.XtraEditors.DateEdit dtFEmision;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private DevExpress.XtraEditors.CheckEdit checkContratado;
		private C1.Data.C1DataSet cdsGuiaPunto;
		private C1.Data.C1DataView cdvGuiaPunto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPuntoLlegada;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPuntoPartida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiesel;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAdelanto;
		private DevExpress.XtraEditors.SimpleButton btnPunto;
		private DevExpress.XtraEditors.SimpleButton btnPunto2;
		private System.Windows.Forms.Label lblAnulado;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCarro1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCarro2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCarro3;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private C1.Data.C1DataSet cdsPersonal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdGuiaRemisionG;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRucCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreCli;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtChofer;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRucChofer;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTonPedido;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPlaca;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtComAut;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtComCons;
		private C1.Data.C1DataView cdvCliente;
		private C1.Data.C1DataSet cdsCliente;

		private C1.Data.C1DataRow drGuia;
	//	private C1.Data.C1DataRow drCliente;
		int nroguia=0;
		private System.Windows.Forms.Label label26;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTextoTonelaje;
		private System.Windows.Forms.Label label27;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbChofer;
		private Infragistics.Win.Misc.UltraButton btCliente;
		private DevExpress.XtraEditors.SimpleButton btDigitalizacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdChofer;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label29;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreCliD;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRucClienteD;
		int editar=0;

		public GuiaRemisionG()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_GuiaRemisionDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nroBultos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("volumen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("contenido");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionPunto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("GuiaRemisionPunto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGuiaRemisionPunto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(GuiaRemisionG));
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			this.cdsGuiaRemisionG = new C1.Data.C1DataSet();
			this.cdvGuiaRemisionG = new C1.Data.C1DataView();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.label6 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.chkVenta = new DevExpress.XtraEditors.CheckEdit();
			this.checkCompra = new DevExpress.XtraEditors.CheckEdit();
			this.checkTransformacion = new DevExpress.XtraEditors.CheckEdit();
			this.checkConsignacion = new DevExpress.XtraEditors.CheckEdit();
			this.checkTraslado = new DevExpress.XtraEditors.CheckEdit();
			this.checkDevolucion = new DevExpress.XtraEditors.CheckEdit();
			this.checkImportacion = new DevExpress.XtraEditors.CheckEdit();
			this.checkExportacion = new DevExpress.XtraEditors.CheckEdit();
			this.checkOtros = new DevExpress.XtraEditors.CheckEdit();
			this.dtFinicioT = new DevExpress.XtraEditors.DateEdit();
			this.dtFTermT = new DevExpress.XtraEditors.DateEdit();
			this.dtFEmision = new DevExpress.XtraEditors.DateEdit();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtRucCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtChofer = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.checkContratado = new DevExpress.XtraEditors.CheckEdit();
			this.cdsGuiaPunto = new C1.Data.C1DataSet();
			this.cdvGuiaPunto = new C1.Data.C1DataView();
			this.cmbPuntoLlegada = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbPuntoPartida = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtDiesel = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtAdelanto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnPunto = new DevExpress.XtraEditors.SimpleButton();
			this.btnPunto2 = new DevExpress.XtraEditors.SimpleButton();
			this.lblAnulado = new System.Windows.Forms.Label();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cmbCarro1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCarro2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCarro3 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNombreCli = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.txtObservacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdGuiaRemisionG = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtRucChofer = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.txtTonPedido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label21 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPlaca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.label24 = new System.Windows.Forms.Label();
			this.txtComAut = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtComCons = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdvCliente = new C1.Data.C1DataView();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.label26 = new System.Windows.Forms.Label();
			this.txtTextoTonelaje = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label27 = new System.Windows.Forms.Label();
			this.cmbChofer = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.btDigitalizacion = new DevExpress.XtraEditors.SimpleButton();
			this.txtIdChofer = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label28 = new System.Windows.Forms.Label();
			this.label29 = new System.Windows.Forms.Label();
			this.txtNombreCliD = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRucClienteD = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRemisionG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaRemisionG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkVenta.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkCompra.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkTransformacion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkConsignacion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkTraslado.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkDevolucion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkImportacion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkExportacion.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkOtros.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFinicioT.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFTermT.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFEmision.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtChofer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.checkContratado.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaPunto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaPunto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoLlegada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoPartida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiesel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdelanto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCarro1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCarro2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCarro3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreCli)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGuiaRemisionG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucChofer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTonPedido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComAut)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComCons)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTextoTonelaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbChofer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdChofer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreCliD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucClienteD)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsGuiaRemisionG
			// 
			this.cdsGuiaRemisionG.BindingContextCtrl = this;
			this.cdsGuiaRemisionG.DataLibrary = "LibFacturacion";
			this.cdsGuiaRemisionG.DataLibraryUrl = "";
			this.cdsGuiaRemisionG.DataSetDef = "dsGuiaRemisionG";
			this.cdsGuiaRemisionG.EnforceConstraints = false;
			this.cdsGuiaRemisionG.FillOnRequest = false;
			this.cdsGuiaRemisionG.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGuiaRemisionG.Name = "cdsGuiaRemisionG";
			this.cdsGuiaRemisionG.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGuiaRemisionG.SchemaDef = null;
			this.cdsGuiaRemisionG.BeforeFill += new C1.Data.FillEventHandler(this.cdsGuiaRemisionG_BeforeFill);
			// 
			// cdvGuiaRemisionG
			// 
			this.cdvGuiaRemisionG.BindingContextCtrl = this;
			this.cdvGuiaRemisionG.DataSet = this.cdsGuiaRemisionG;
			this.cdvGuiaRemisionG.TableName = "";
			this.cdvGuiaRemisionG.TableViewName = "GuiaRemisionG";
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_GuiaRemisionG";
			this.barraDato1.DataNombreId = "idGuiaRemisionG";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsGuiaRemisionG;
			this.barraDato1.DataTabla = "GuiaRemisionG";
			this.barraDato1.DataTablaHija = "GuiaRemisionDetalle";
			this.barraDato1.Location = new System.Drawing.Point(240, 569);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 25);
			this.barraDato1.TabIndex = 13;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(496, 43);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(75, 16);
			this.label6.TabIndex = 38;
			this.label6.Text = "Guia Remision";
			// 
			// txtCodigo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance1;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.NroGuia"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtCodigo.Location = new System.Drawing.Point(576, 43);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(120, 22);
			this.txtCodigo.TabIndex = 37;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_GuiaRemisionG.GuiaRemisionG - guiaRemisionDetalle";
			this.ultraGrid1.DataSource = this.cdsGuiaRemisionG;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn3.PromptChar = ' ';
			ultraGridColumn3.Width = 132;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 144;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn5.PromptChar = ' ';
			ultraGridColumn5.Width = 113;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn6.PromptChar = ' ';
			ultraGridColumn6.Width = 111;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 175;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(32, 431);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(696, 138);
			this.ultraGrid1.TabIndex = 39;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(272, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(56, 16);
			this.label1.TabIndex = 40;
			this.label1.Text = "F. Emision";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(24, 69);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(93, 16);
			this.label2.TabIndex = 41;
			this.label2.Text = "F. Term. Traslado";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(90, 16);
			this.label3.TabIndex = 42;
			this.label3.Text = "F. Inicio Traslado";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(400, 146);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(94, 16);
			this.label5.TabIndex = 44;
			this.label5.Text = "Punto de LLegada";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(120, 176);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(120, 16);
			this.label7.TabIndex = 45;
			this.label7.Text = "Nombre o Razon Social";
			this.label7.Visible = false;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(592, 176);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(33, 16);
			this.label8.TabIndex = 46;
			this.label8.Text = "RUC :";
			this.label8.Visible = false;
			// 
			// chkVenta
			// 
			this.chkVenta.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.venta"));
			this.chkVenta.EditValue = true;
			this.chkVenta.Location = new System.Drawing.Point(8, 121);
			this.chkVenta.Name = "chkVenta";
			// 
			// chkVenta.Properties
			// 
			this.chkVenta.Properties.Caption = "Venta";
			this.chkVenta.Size = new System.Drawing.Size(56, 18);
			this.chkVenta.TabIndex = 47;
			// 
			// checkCompra
			// 
			this.checkCompra.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.compra"));
			this.checkCompra.EditValue = true;
			this.checkCompra.Location = new System.Drawing.Point(72, 121);
			this.checkCompra.Name = "checkCompra";
			// 
			// checkCompra.Properties
			// 
			this.checkCompra.Properties.Caption = "Compra";
			this.checkCompra.Size = new System.Drawing.Size(64, 18);
			this.checkCompra.TabIndex = 48;
			// 
			// checkTransformacion
			// 
			this.checkTransformacion.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.transformacion"));
			this.checkTransformacion.EditValue = true;
			this.checkTransformacion.Location = new System.Drawing.Point(136, 121);
			this.checkTransformacion.Name = "checkTransformacion";
			// 
			// checkTransformacion.Properties
			// 
			this.checkTransformacion.Properties.Caption = "Transformacion";
			this.checkTransformacion.Size = new System.Drawing.Size(104, 18);
			this.checkTransformacion.TabIndex = 49;
			// 
			// checkConsignacion
			// 
			this.checkConsignacion.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.transformacion"));
			this.checkConsignacion.EditValue = true;
			this.checkConsignacion.Location = new System.Drawing.Point(240, 121);
			this.checkConsignacion.Name = "checkConsignacion";
			// 
			// checkConsignacion.Properties
			// 
			this.checkConsignacion.Properties.Caption = "Consignacion";
			this.checkConsignacion.Size = new System.Drawing.Size(96, 18);
			this.checkConsignacion.TabIndex = 50;
			// 
			// checkTraslado
			// 
			this.checkTraslado.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.traslado"));
			this.checkTraslado.EditValue = true;
			this.checkTraslado.Location = new System.Drawing.Point(344, 121);
			this.checkTraslado.Name = "checkTraslado";
			// 
			// checkTraslado.Properties
			// 
			this.checkTraslado.Properties.Caption = "Traslado";
			this.checkTraslado.Size = new System.Drawing.Size(72, 18);
			this.checkTraslado.TabIndex = 51;
			// 
			// checkDevolucion
			// 
			this.checkDevolucion.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.devolucion"));
			this.checkDevolucion.EditValue = true;
			this.checkDevolucion.Location = new System.Drawing.Point(424, 121);
			this.checkDevolucion.Name = "checkDevolucion";
			// 
			// checkDevolucion.Properties
			// 
			this.checkDevolucion.Properties.Caption = "Devolucion";
			this.checkDevolucion.Size = new System.Drawing.Size(80, 18);
			this.checkDevolucion.TabIndex = 52;
			// 
			// checkImportacion
			// 
			this.checkImportacion.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.importacion"));
			this.checkImportacion.EditValue = true;
			this.checkImportacion.Location = new System.Drawing.Point(504, 121);
			this.checkImportacion.Name = "checkImportacion";
			// 
			// checkImportacion.Properties
			// 
			this.checkImportacion.Properties.Caption = "Importacion";
			this.checkImportacion.Size = new System.Drawing.Size(88, 18);
			this.checkImportacion.TabIndex = 53;
			// 
			// checkExportacion
			// 
			this.checkExportacion.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.exportacion"));
			this.checkExportacion.EditValue = true;
			this.checkExportacion.Location = new System.Drawing.Point(592, 121);
			this.checkExportacion.Name = "checkExportacion";
			// 
			// checkExportacion.Properties
			// 
			this.checkExportacion.Properties.Caption = "Exportacion";
			this.checkExportacion.Size = new System.Drawing.Size(88, 18);
			this.checkExportacion.TabIndex = 54;
			// 
			// checkOtros
			// 
			this.checkOtros.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.otros"));
			this.checkOtros.EditValue = true;
			this.checkOtros.Location = new System.Drawing.Point(680, 121);
			this.checkOtros.Name = "checkOtros";
			// 
			// checkOtros.Properties
			// 
			this.checkOtros.Properties.Caption = "Otros";
			this.checkOtros.Size = new System.Drawing.Size(56, 18);
			this.checkOtros.TabIndex = 55;
			// 
			// dtFinicioT
			// 
			this.dtFinicioT.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.fInicioTraslado"));
			this.dtFinicioT.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFinicioT.Location = new System.Drawing.Point(128, 43);
			this.dtFinicioT.Name = "dtFinicioT";
			// 
			// dtFinicioT.Properties
			// 
			this.dtFinicioT.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFinicioT.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFinicioT.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFinicioT.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFinicioT.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFinicioT.Size = new System.Drawing.Size(96, 20);
			this.dtFinicioT.TabIndex = 56;
			// 
			// dtFTermT
			// 
			this.dtFTermT.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.fTermTraslado"));
			this.dtFTermT.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFTermT.Location = new System.Drawing.Point(128, 69);
			this.dtFTermT.Name = "dtFTermT";
			// 
			// dtFTermT.Properties
			// 
			this.dtFTermT.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFTermT.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFTermT.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFTermT.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFTermT.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFTermT.Size = new System.Drawing.Size(96, 20);
			this.dtFTermT.TabIndex = 57;
			// 
			// dtFEmision
			// 
			this.dtFEmision.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.fEmision2"));
			this.dtFEmision.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFEmision.Location = new System.Drawing.Point(352, 43);
			this.dtFEmision.Name = "dtFEmision";
			// 
			// dtFEmision.Properties
			// 
			this.dtFEmision.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFEmision.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFEmision.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFEmision.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFEmision.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFEmision.Size = new System.Drawing.Size(96, 20);
			this.dtFEmision.TabIndex = 58;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(8, 103);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(168, 18);
			this.label9.TabIndex = 59;
			this.label9.Text = "MOTIVO DEL TRASLADO";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(8, 172);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(59, 18);
			this.label10.TabIndex = 60;
			this.label10.Text = "ORIGEN";
			this.label10.Visible = false;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(160, 232);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(451, 18);
			this.label11.TabIndex = 61;
			this.label11.Text = "IDENTIFICACION DE LA PERSONA ENCARGADA DEL TRANSPORTE";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 256);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(37, 16);
			this.label4.TabIndex = 62;
			this.label4.Text = "Chofer";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(368, 280);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(33, 16);
			this.label12.TabIndex = 63;
			this.label12.Text = "RUC :";
			// 
			// txtRucCliente
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRucCliente.Appearance = appearance7;
			this.txtRucCliente.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsGuiaRemisionG, "_GuiaRemisionG.rucCliente"));
			this.txtRucCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtRucCliente.Location = new System.Drawing.Point(632, 176);
			this.txtRucCliente.Name = "txtRucCliente";
			this.txtRucCliente.Size = new System.Drawing.Size(96, 22);
			this.txtRucCliente.TabIndex = 65;
			this.txtRucCliente.Visible = false;
			// 
			// txtChofer
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtChofer.Appearance = appearance8;
			this.txtChofer.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsGuiaRemisionG, "_GuiaRemisionG.nombreTrasnp"));
			this.txtChofer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtChofer.Location = new System.Drawing.Point(64, 280);
			this.txtChofer.Name = "txtChofer";
			this.txtChofer.Size = new System.Drawing.Size(272, 22);
			this.txtChofer.TabIndex = 67;
			this.txtChofer.Visible = false;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(8, 146);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(87, 16);
			this.label13.TabIndex = 69;
			this.label13.Text = "Punto de Partida";
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(512, 280);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(42, 16);
			this.label14.TabIndex = 71;
			this.label14.Text = "Diesel :";
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(648, 280);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(56, 16);
			this.label15.TabIndex = 72;
			this.label15.Text = "Adelanto :";
			// 
			// checkContratado
			// 
			this.checkContratado.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsGuiaRemisionG, "GuiaRemisionG.contratado"));
			this.checkContratado.EditValue = true;
			this.checkContratado.Location = new System.Drawing.Point(16, 9);
			this.checkContratado.Name = "checkContratado";
			// 
			// checkContratado.Properties
			// 
			this.checkContratado.Properties.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.checkContratado.Properties.Appearance.Options.UseFont = true;
			this.checkContratado.Properties.Caption = "Contratado";
			this.checkContratado.Size = new System.Drawing.Size(112, 21);
			this.checkContratado.TabIndex = 75;
			this.checkContratado.Validated += new System.EventHandler(this.checkContratado_Validated);
			this.checkContratado.CheckedChanged += new System.EventHandler(this.checkContratado_CheckedChanged);
			// 
			// cdsGuiaPunto
			// 
			this.cdsGuiaPunto.BindingContextCtrl = this;
			this.cdsGuiaPunto.DataLibrary = "LibFacturacion";
			this.cdsGuiaPunto.DataLibraryUrl = "";
			this.cdsGuiaPunto.DataSetDef = "dsGuiaRemisionPunto";
			this.cdsGuiaPunto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGuiaPunto.Name = "cdsGuiaPunto";
			this.cdsGuiaPunto.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGuiaPunto.SchemaDef = null;
			this.cdsGuiaPunto.BeforeFill += new C1.Data.FillEventHandler(this.cdsGuiaPunto_BeforeFill);
			// 
			// cdvGuiaPunto
			// 
			this.cdvGuiaPunto.BindingContextCtrl = this;
			this.cdvGuiaPunto.DataSet = this.cdsGuiaPunto;
			this.cdvGuiaPunto.Sort = "nombre";
			this.cdvGuiaPunto.TableName = "";
			this.cdvGuiaPunto.TableViewName = "GuiaRemisionPunto";
			// 
			// cmbPuntoLlegada
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPuntoLlegada.Appearance = appearance9;
			this.cmbPuntoLlegada.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPuntoLlegada.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPuntoLlegada.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.idPuntoLlegada"));
			this.cmbPuntoLlegada.DataSource = this.cdvGuiaPunto;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbPuntoLlegada.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPuntoLlegada.DisplayMember = "Nombre";
			this.cmbPuntoLlegada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPuntoLlegada.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPuntoLlegada.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPuntoLlegada.Location = new System.Drawing.Point(504, 146);
			this.cmbPuntoLlegada.Name = "cmbPuntoLlegada";
			this.cmbPuntoLlegada.Size = new System.Drawing.Size(168, 21);
			this.cmbPuntoLlegada.TabIndex = 76;
			this.cmbPuntoLlegada.ValueMember = "idGuiaRemisionPunto";
			// 
			// cmbPuntoPartida
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPuntoPartida.Appearance = appearance10;
			this.cmbPuntoPartida.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPuntoPartida.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPuntoPartida.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.idPuntoPartida"));
			this.cmbPuntoPartida.DataSource = this.cdvGuiaPunto;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbPuntoPartida.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPuntoPartida.DisplayMember = "Nombre";
			this.cmbPuntoPartida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPuntoPartida.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPuntoPartida.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPuntoPartida.Location = new System.Drawing.Point(104, 146);
			this.cmbPuntoPartida.Name = "cmbPuntoPartida";
			this.cmbPuntoPartida.Size = new System.Drawing.Size(160, 21);
			this.cmbPuntoPartida.TabIndex = 77;
			this.cmbPuntoPartida.ValueMember = "idGuiaRemisionPunto";
			// 
			// txtDiesel
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiesel.Appearance = appearance11;
			this.txtDiesel.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "GuiaRemisionG.diesel"));
			this.txtDiesel.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtDiesel.FormatString = "#,##0.00";
			this.txtDiesel.Location = new System.Drawing.Point(560, 280);
			this.txtDiesel.Name = "txtDiesel";
			this.txtDiesel.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDiesel.PromptChar = ' ';
			this.txtDiesel.Size = new System.Drawing.Size(72, 22);
			this.txtDiesel.TabIndex = 108;
			// 
			// txtAdelanto
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAdelanto.Appearance = appearance12;
			this.txtAdelanto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.adelanto"));
			this.txtAdelanto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtAdelanto.FormatString = "#,##0.00";
			this.txtAdelanto.Location = new System.Drawing.Point(704, 280);
			this.txtAdelanto.Name = "txtAdelanto";
			this.txtAdelanto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAdelanto.PromptChar = ' ';
			this.txtAdelanto.Size = new System.Drawing.Size(64, 22);
			this.txtAdelanto.TabIndex = 109;
			// 
			// btnPunto
			// 
			this.btnPunto.Image = ((System.Drawing.Image)(resources.GetObject("btnPunto.Image")));
			this.btnPunto.Location = new System.Drawing.Point(680, 146);
			this.btnPunto.Name = "btnPunto";
			this.btnPunto.Size = new System.Drawing.Size(24, 25);
			this.btnPunto.TabIndex = 110;
			this.btnPunto.ToolTip = "Retenciones de artículos";
			this.btnPunto.Click += new System.EventHandler(this.btnPunto_Click);
			// 
			// btnPunto2
			// 
			this.btnPunto2.Image = ((System.Drawing.Image)(resources.GetObject("btnPunto2.Image")));
			this.btnPunto2.Location = new System.Drawing.Point(272, 146);
			this.btnPunto2.Name = "btnPunto2";
			this.btnPunto2.Size = new System.Drawing.Size(24, 25);
			this.btnPunto2.TabIndex = 111;
			this.btnPunto2.ToolTip = "Retenciones de artículos";
			this.btnPunto2.Click += new System.EventHandler(this.btnPunto2_Click);
			// 
			// lblAnulado
			// 
			this.lblAnulado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblAnulado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblAnulado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblAnulado.Location = new System.Drawing.Point(624, 577);
			this.lblAnulado.Name = "lblAnulado";
			this.lblAnulado.Size = new System.Drawing.Size(112, 25);
			this.lblAnulado.TabIndex = 114;
			this.lblAnulado.Text = "ANULADO";
			this.lblAnulado.Visible = false;
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
			// cmbCarro1
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCarro1.Appearance = appearance13;
			this.cmbCarro1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCarro1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCarro1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.idCarro1"));
			this.cmbCarro1.DataSource = this.cdvProyecto;
			ultraGridBand4.AddButtonCaption = "GuiaRemisionPunto";
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn19.Header.VisiblePosition = 7;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 8;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 9;
			ultraGridBand4.Columns.AddRange(new object[] {
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
			this.cmbCarro1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCarro1.DisplayMember = "Nombre";
			this.cmbCarro1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCarro1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCarro1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCarro1.Location = new System.Drawing.Point(136, 345);
			this.cmbCarro1.Name = "cmbCarro1";
			this.cmbCarro1.Size = new System.Drawing.Size(192, 21);
			this.cmbCarro1.TabIndex = 119;
			this.cmbCarro1.ValueMember = "idProyecto";
			this.cmbCarro1.Validated += new System.EventHandler(this.cmbCarro1_Validated);
			this.cmbCarro1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCarro1_MouseUp);
			// 
			// cmbCarro2
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCarro2.Appearance = appearance14;
			this.cmbCarro2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCarro2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCarro2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.idCarro2"));
			this.cmbCarro2.DataSource = this.cdvProyecto;
			ultraGridBand5.AddButtonCaption = "GuiaRemisionPunto";
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.LockedWidth = true;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.LockedWidth = true;
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn24.LockedWidth = true;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn25.LockedWidth = true;
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.LockedWidth = true;
			ultraGridColumn27.Header.VisiblePosition = 6;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.LockedWidth = true;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.LockedWidth = true;
			ultraGridColumn29.Header.VisiblePosition = 7;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.LockedWidth = true;
			ultraGridColumn30.Header.VisiblePosition = 8;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.LockedWidth = true;
			ultraGridColumn31.Header.VisiblePosition = 9;
			ultraGridBand5.Columns.AddRange(new object[] {
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
			this.cmbCarro2.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCarro2.DisplayMember = "Nombre";
			this.cmbCarro2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCarro2.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCarro2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCarro2.Location = new System.Drawing.Point(344, 345);
			this.cmbCarro2.Name = "cmbCarro2";
			this.cmbCarro2.Size = new System.Drawing.Size(192, 21);
			this.cmbCarro2.TabIndex = 120;
			this.cmbCarro2.ValueMember = "idProyecto";
			this.cmbCarro2.Validated += new System.EventHandler(this.cmbCarro2_Validated);
			this.cmbCarro2.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCarro2_MouseUp);
			// 
			// cmbCarro3
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCarro3.Appearance = appearance15;
			this.cmbCarro3.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCarro3.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCarro3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.idCarro3"));
			this.cmbCarro3.DataSource = this.cdvProyecto;
			ultraGridBand6.AddButtonCaption = "GuiaRemisionPunto";
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 3;
			ultraGridColumn35.Header.VisiblePosition = 4;
			ultraGridColumn36.Header.VisiblePosition = 5;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 6;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 2;
			ultraGridColumn39.Header.VisiblePosition = 7;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 8;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 9;
			ultraGridBand6.Columns.AddRange(new object[] {
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
			this.cmbCarro3.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCarro3.DisplayMember = "Nombre";
			this.cmbCarro3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCarro3.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCarro3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCarro3.Location = new System.Drawing.Point(552, 345);
			this.cmbCarro3.Name = "cmbCarro3";
			this.cmbCarro3.Size = new System.Drawing.Size(176, 21);
			this.cmbCarro3.TabIndex = 121;
			this.cmbCarro3.ValueMember = "idProyecto";
			this.cmbCarro3.Validated += new System.EventHandler(this.cmbCarro3_Validated);
			this.cmbCarro3.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCarro3_MouseUp);
			// 
			// txtNombreCli
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreCli.Appearance = appearance16;
			this.txtNombreCli.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.NombreCliente"));
			this.txtNombreCli.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreCli.Location = new System.Drawing.Point(248, 176);
			this.txtNombreCli.Name = "txtNombreCli";
			this.txtNombreCli.Size = new System.Drawing.Size(328, 22);
			this.txtNombreCli.TabIndex = 122;
			this.txtNombreCli.Visible = false;
			this.txtNombreCli.Validating += new System.ComponentModel.CancelEventHandler(this.txtNombre_Validating);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(200, 327);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(56, 16);
			this.label17.TabIndex = 123;
			this.label17.Text = "Maquina 1";
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(408, 327);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(56, 16);
			this.label18.TabIndex = 124;
			this.label18.Text = "Maquina 2";
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(616, 327);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(56, 16);
			this.label19.TabIndex = 125;
			this.label19.Text = "Maquina 3";
			// 
			// cmbPersonal
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPersonal.Appearance = appearance17;
			this.cmbPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.idTransportista"));
			this.cmbPersonal.DataSource = this.cdvPersonal;
			ultraGridBand7.AddButtonCaption = "Cliente";
			ultraGridColumn42.Header.VisiblePosition = 0;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 2;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 4;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 5;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 6;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 7;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 8;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 9;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 1;
			ultraGridColumn50.Width = 231;
			ultraGridColumn51.Header.VisiblePosition = 10;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 3;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 11;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 12;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 13;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 14;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 15;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 16;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 17;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 18;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 19;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 20;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 21;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 22;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 23;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 24;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 25;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 26;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 27;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 28;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 29;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 30;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 31;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 32;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 33;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 34;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 35;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 36;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 37;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 38;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 39;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 40;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 41;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 42;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 43;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 44;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 45;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 46;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 47;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 48;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 49;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 50;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 51;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 52;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 53;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 54;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 55;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 56;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 57;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 58;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 59;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 60;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 61;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 62;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 63;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn106.Header.VisiblePosition = 64;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 65;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 66;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.VisiblePosition = 67;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 68;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 69;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 70;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 71;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 72;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 73;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 74;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 75;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 76;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 77;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 78;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 79;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 80;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 81;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 82;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 83;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 84;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.VisiblePosition = 85;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn128.Header.VisiblePosition = 86;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.VisiblePosition = 87;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn130.Header.VisiblePosition = 88;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn131.Header.VisiblePosition = 89;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn132.Header.VisiblePosition = 90;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn133.Header.VisiblePosition = 91;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn134.Header.VisiblePosition = 92;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn135.Header.VisiblePosition = 93;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn136.Header.VisiblePosition = 94;
			ultraGridColumn136.Hidden = true;
			ultraGridBand7.Columns.AddRange(new object[] {
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
																										 ultraGridColumn136});
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPersonal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonal.Location = new System.Drawing.Point(64, 256);
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(272, 21);
			this.cmbPersonal.TabIndex = 126;
			this.cmbPersonal.ValueMember = "idPersonal";
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
			// txtObservacion
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservacion.Appearance = appearance18;
			this.txtObservacion.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsGuiaRemisionG, "_GuiaRemisionG.observaciones"));
			this.txtObservacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtObservacion.Location = new System.Drawing.Point(48, 396);
			this.txtObservacion.Name = "txtObservacion";
			this.txtObservacion.Size = new System.Drawing.Size(680, 22);
			this.txtObservacion.TabIndex = 127;
			// 
			// txtIdGuiaRemisionG
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdGuiaRemisionG.Appearance = appearance19;
			this.txtIdGuiaRemisionG.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.idGuiaRemisionG"));
			this.txtIdGuiaRemisionG.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtIdGuiaRemisionG.Location = new System.Drawing.Point(144, 9);
			this.txtIdGuiaRemisionG.Name = "txtIdGuiaRemisionG";
			this.txtIdGuiaRemisionG.Size = new System.Drawing.Size(96, 22);
			this.txtIdGuiaRemisionG.TabIndex = 128;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// txtRucChofer
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRucChofer.Appearance = appearance20;
			this.txtRucChofer.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsGuiaRemisionG, "_GuiaRemisionG.rucTransp"));
			this.txtRucChofer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtRucChofer.Location = new System.Drawing.Point(408, 280);
			this.txtRucChofer.Name = "txtRucChofer";
			this.txtRucChofer.Size = new System.Drawing.Size(96, 22);
			this.txtRucChofer.TabIndex = 129;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Location = new System.Drawing.Point(24, 327);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(101, 16);
			this.label20.TabIndex = 130;
			this.label20.Text = "Ton. Pedido Cliente";
			// 
			// txtTonPedido
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTonPedido.Appearance = appearance21;
			this.txtTonPedido.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "GuiaRemisionG.TonelajePedido"));
			this.txtTonPedido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTonPedido.FormatString = "#,##0.00";
			this.txtTonPedido.Location = new System.Drawing.Point(48, 345);
			this.txtTonPedido.Name = "txtTonPedido";
			this.txtTonPedido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTonPedido.PromptChar = ' ';
			this.txtTonPedido.Size = new System.Drawing.Size(52, 22);
			this.txtTonPedido.TabIndex = 131;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.Location = new System.Drawing.Point(512, 304);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(46, 16);
			this.label21.TabIndex = 132;
			this.label21.Text = "Celular :";
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.Location = new System.Drawing.Point(664, 304);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(38, 16);
			this.label22.TabIndex = 133;
			this.label22.Text = "Placa :";
			// 
			// txtCelular
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance22;
			this.txtCelular.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsGuiaRemisionG, "_GuiaRemisionG.Celular"));
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtCelular.Location = new System.Drawing.Point(560, 304);
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(96, 22);
			this.txtCelular.TabIndex = 134;
			// 
			// txtPlaca
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPlaca.Appearance = appearance23;
			this.txtPlaca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsGuiaRemisionG, "_GuiaRemisionG.Placa"));
			this.txtPlaca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtPlaca.Location = new System.Drawing.Point(712, 304);
			this.txtPlaca.Name = "txtPlaca";
			this.txtPlaca.Size = new System.Drawing.Size(64, 22);
			this.txtPlaca.TabIndex = 135;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.Location = new System.Drawing.Point(424, 256);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(93, 16);
			this.label23.TabIndex = 136;
			this.label23.Text = "Comb. Autorizado";
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.Location = new System.Drawing.Point(592, 256);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(95, 16);
			this.label24.TabIndex = 137;
			this.label24.Text = "Comb. Consumido";
			// 
			// txtComAut
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComAut.Appearance = appearance24;
			this.txtComAut.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "GuiaRemisionG.CombAutorizado"));
			this.txtComAut.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComAut.FormatString = "#,##0.00";
			this.txtComAut.ImageTransparentColor = System.Drawing.Color.Black;
			this.txtComAut.Location = new System.Drawing.Point(520, 256);
			this.txtComAut.Name = "txtComAut";
			this.txtComAut.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtComAut.PromptChar = ' ';
			this.txtComAut.Size = new System.Drawing.Size(72, 22);
			this.txtComAut.TabIndex = 138;
			// 
			// txtComCons
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComCons.Appearance = appearance25;
			this.txtComCons.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "GuiaRemisionG.CombConsumido"));
			this.txtComCons.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComCons.FormatString = "#,##0.00";
			this.txtComCons.Location = new System.Drawing.Point(704, 256);
			this.txtComCons.Name = "txtComCons";
			this.txtComCons.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtComCons.PromptChar = ' ';
			this.txtComCons.Size = new System.Drawing.Size(72, 22);
			this.txtComCons.TabIndex = 139;
			// 
			// cdvCliente
			// 
			this.cdvCliente.BindingContextCtrl = this;
			this.cdvCliente.DataSet = this.cdsCliente;
			this.cdvCliente.RowFilter = "proveedor=1  or ambos=1";
			this.cdvCliente.Sort = "Nombre";
			this.cdvCliente.TableName = "";
			this.cdvCliente.TableViewName = "Cliente";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.CaseSensitive = false;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.Location = new System.Drawing.Point(16, 280);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(37, 16);
			this.label26.TabIndex = 144;
			this.label26.Text = "Chofer";
			// 
			// txtTextoTonelaje
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTextoTonelaje.Appearance = appearance26;
			this.txtTextoTonelaje.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.TextoTonelaje"));
			this.txtTextoTonelaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtTextoTonelaje.Location = new System.Drawing.Point(48, 370);
			this.txtTextoTonelaje.Name = "txtTextoTonelaje";
			this.txtTextoTonelaje.Size = new System.Drawing.Size(280, 22);
			this.txtTextoTonelaje.TabIndex = 145;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.Location = new System.Drawing.Point(8, 370);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(32, 16);
			this.label27.TabIndex = 146;
			this.label27.Text = "Texto";
			// 
			// cmbChofer
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbChofer.Appearance = appearance27;
			this.cmbChofer.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbChofer.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbChofer.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.idTransportista"));
			this.cmbChofer.DisplayMember = "";
			this.cmbChofer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbChofer.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbChofer.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbChofer.Location = new System.Drawing.Point(64, 288);
			this.cmbChofer.Name = "cmbChofer";
			this.cmbChofer.Size = new System.Drawing.Size(272, 21);
			this.cmbChofer.TabIndex = 147;
			this.cmbChofer.ValueMember = "";
			this.cmbChofer.ValueChanged += new System.EventHandler(this.cmbChofer_ValueChanged);
			// 
			// btCliente
			// 
			this.btCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCliente.Location = new System.Drawing.Point(336, 288);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(24, 24);
			this.btCliente.TabIndex = 148;
			this.btCliente.Text = "...";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// btDigitalizacion
			// 
			this.btDigitalizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btDigitalizacion.ImageIndex = 0;
			this.btDigitalizacion.Location = new System.Drawing.Point(768, 264);
			this.btDigitalizacion.Name = "btDigitalizacion";
			this.btDigitalizacion.Size = new System.Drawing.Size(24, 25);
			this.btDigitalizacion.TabIndex = 149;
			this.btDigitalizacion.Text = "...";
			this.btDigitalizacion.Click += new System.EventHandler(this.btDigitalizacion_Click);
			// 
			// txtIdChofer
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdChofer.Appearance = appearance28;
			this.txtIdChofer.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "GuiaRemisionG.idTransportista"));
			this.txtIdChofer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdChofer.FormatString = "#,##0.00";
			this.txtIdChofer.ImageTransparentColor = System.Drawing.Color.Black;
			this.txtIdChofer.Location = new System.Drawing.Point(120, 312);
			this.txtIdChofer.Name = "txtIdChofer";
			this.txtIdChofer.PromptChar = ' ';
			this.txtIdChofer.Size = new System.Drawing.Size(72, 22);
			this.txtIdChofer.TabIndex = 150;
			this.txtIdChofer.Visible = false;
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.Location = new System.Drawing.Point(592, 208);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(33, 16);
			this.label28.TabIndex = 156;
			this.label28.Text = "RUC :";
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.Location = new System.Drawing.Point(120, 208);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(120, 16);
			this.label29.TabIndex = 155;
			this.label29.Text = "Nombre o Razon Social";
			// 
			// txtNombreCliD
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreCliD.Appearance = appearance29;
			this.txtNombreCliD.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRemisionG, "_GuiaRemisionG.NombreClienteD"));
			this.txtNombreCliD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreCliD.Location = new System.Drawing.Point(248, 208);
			this.txtNombreCliD.Name = "txtNombreCliD";
			this.txtNombreCliD.Size = new System.Drawing.Size(328, 22);
			this.txtNombreCliD.TabIndex = 158;
			this.txtNombreCliD.Validating += new System.ComponentModel.CancelEventHandler(this.txtNombreCliD_Validating);
			// 
			// txtRucClienteD
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRucClienteD.Appearance = appearance30;
			this.txtRucClienteD.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsGuiaRemisionG, "_GuiaRemisionG.RucClienteD"));
			this.txtRucClienteD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.OfficeXP;
			this.txtRucClienteD.Location = new System.Drawing.Point(632, 208);
			this.txtRucClienteD.Name = "txtRucClienteD";
			this.txtRucClienteD.Size = new System.Drawing.Size(96, 22);
			this.txtRucClienteD.TabIndex = 157;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.Location = new System.Drawing.Point(8, 208);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(107, 18);
			this.label16.TabIndex = 159;
			this.label16.Text = "DESTINATARIO";
			// 
			// GuiaRemisionG
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(792, 608);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label28);
			this.Controls.Add(this.label29);
			this.Controls.Add(this.txtNombreCliD);
			this.Controls.Add(this.txtRucClienteD);
			this.Controls.Add(this.txtIdChofer);
			this.Controls.Add(this.btDigitalizacion);
			this.Controls.Add(this.btCliente);
			this.Controls.Add(this.cmbChofer);
			this.Controls.Add(this.label27);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtTextoTonelaje);
			this.Controls.Add(this.txtComCons);
			this.Controls.Add(this.txtComAut);
			this.Controls.Add(this.txtPlaca);
			this.Controls.Add(this.txtCelular);
			this.Controls.Add(this.txtTonPedido);
			this.Controls.Add(this.txtRucChofer);
			this.Controls.Add(this.txtIdGuiaRemisionG);
			this.Controls.Add(this.txtObservacion);
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.txtNombreCli);
			this.Controls.Add(this.cmbCarro3);
			this.Controls.Add(this.cmbCarro2);
			this.Controls.Add(this.cmbCarro1);
			this.Controls.Add(this.lblAnulado);
			this.Controls.Add(this.btnPunto2);
			this.Controls.Add(this.btnPunto);
			this.Controls.Add(this.txtAdelanto);
			this.Controls.Add(this.txtDiesel);
			this.Controls.Add(this.cmbPuntoPartida);
			this.Controls.Add(this.cmbPuntoLlegada);
			this.Controls.Add(this.checkContratado);
			this.Controls.Add(this.txtChofer);
			this.Controls.Add(this.txtRucCliente);
			this.Controls.Add(this.dtFEmision);
			this.Controls.Add(this.dtFTermT);
			this.Controls.Add(this.dtFinicioT);
			this.Controls.Add(this.checkOtros);
			this.Controls.Add(this.checkExportacion);
			this.Controls.Add(this.checkImportacion);
			this.Controls.Add(this.checkDevolucion);
			this.Controls.Add(this.checkTraslado);
			this.Controls.Add(this.checkConsignacion);
			this.Controls.Add(this.checkTransformacion);
			this.Controls.Add(this.checkCompra);
			this.Controls.Add(this.chkVenta);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.barraDato1);
			this.Name = "GuiaRemisionG";
			this.Text = "Gruia de Remision Gruas";
			this.Load += new System.EventHandler(this.GuiaRemisionG_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRemisionG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaRemisionG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkVenta.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkCompra.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkTransformacion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkConsignacion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkTraslado.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkDevolucion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkImportacion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkExportacion.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkOtros.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFinicioT.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFTermT.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFEmision.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtChofer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.checkContratado.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaPunto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiaPunto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoLlegada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPuntoPartida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiesel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdelanto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCarro1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCarro2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCarro3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreCli)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGuiaRemisionG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucChofer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTonPedido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComAut)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComCons)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTextoTonelaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbChofer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdChofer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreCliD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucClienteD)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
	
			barraDato1.CrearRegistro();
			editar=0;
			activar();
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drGuia = cdsGuiaRemisionG.TableViews["GuiaRemisionG"].Rows[0];
				if ((bool)drGuia["Borrar"] == true) this.lblAnulado.Visible = true;
				else this.lblAnulado.Visible = false;
			}
			catch{};
		}
		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
					if (!Verifica()) return;
					barraDato1.GrabaRegistro();
		}

		private void GuiaRemisionG_Load(object sender, System.EventArgs e)
		{
			txtIdGuiaRemisionG.Width = 0;
			this.barraDato1.HabilitaControles(false);  
	  	this.barraDato1.PosUltima(); 
			#region Guias Natranscom
			if (Funcion.bEjecutaSQL(cdsPersonal, "Exec SeteoGExiste 'GNAT'")) 				
			{
				btDigitalizacion.Visible=true;
				label10.Visible=true;
				label7.Visible=true;
				txtNombreCli.Visible=true;
				label8.Visible=true;
				txtRucCliente.Visible=true;
			}
					#endregion
		}
		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro(); 
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro(); 
			nroguia =int.Parse(drGuia["nroGuia"].ToString() );
			editar=1;
			activar();
		}

		private void cdsGuiaRemisionG_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGuiaRemisionG.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnPunto_Click(object sender, System.EventArgs e)
		{
			using (GuiaPunto miGuiaPunto = new GuiaPunto())
			{
				miGuiaPunto.ShowDialog();
			}
			this.cdsGuiaPunto.Clear();  
			this.cdsGuiaPunto.Fill();  
			this.cmbPuntoLlegada.Refresh();   
			this.cmbPuntoPartida.Refresh();

		}

		private void btnPunto2_Click(object sender, System.EventArgs e)
		{

			using (GuiaPunto miGuiaPunto = new GuiaPunto())
			{
				miGuiaPunto.ShowDialog();
			}
			this.cdsGuiaPunto.Clear();  
			this.cdsGuiaPunto.Fill();  
			this.cmbPuntoLlegada.Refresh();   
			this.cmbPuntoPartida.Refresh();


		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			try
			{	
				switch(e.KeyValue)
				{

					case (int) Keys.Escape:
						this.Close();
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
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;
									
					case (int) Keys.Down:
					
						if (ultraGrid1.ActiveCell.DroppedDown) return;
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						//					MessageBox.Show("baja a la siguiente : valor : " + ultraGrid1.ActiveCell.Value.ToString());  
						
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;
					
					case (int) Keys.Enter:
						
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;
				}
			}
			catch(Exception ex)
			{
				Console.Write(ex.Message);  
			}

		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			drGuia = cdsGuiaRemisionG.TableViews["GuiaRemisionG"].Rows[0];

			if ((bool) drGuia["Borrar"] == false)
			{
				if (DialogResult.No == MessageBox.Show("¿Desea Anular la Guia de Remision ?", "Confirmación", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)) 
					return;
				this.lblAnulado.Visible = true;
				drGuia["Borrar"] = true;
				this.cdsGuiaRemisionG.Update();
			}
			else
			{
				if (DialogResult.No == MessageBox.Show("¿Desea Borrar la Guia de Remision ?\n" +
					"Este proceso NO puede ser Revertido.", "Confirmación", 
					MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2) ) return;
				string stBorrar = "Delete GuiaRemisiondetalle Where idGuiaRemisionG = '" + txtIdGuiaRemisionG.Text.ToString()  + "'" ;
				Funcion.EjecutaSQL(cdsGuiaRemisionG, stBorrar, true);
				this.barraDato1.BorraRegistro();
			}

		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
			if (!(bool) this.drGuia["Borrar"])
			{
				MessageBox.Show("Acción solo es permitida cuando el registro esta anulado",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea Restaurar la Guia de Remision Anulada?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question ))
				return;
			drGuia = cdsGuiaRemisionG.TableViews["GuiaRemisionG"].Rows[0];
			this.drGuia["Borrar"] = false;
			this.lblAnulado.Visible = false;
			this.cdsGuiaRemisionG.Update();
		}

		private void txtNombre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stFiltro = "0";
			string stNombre = txtNombreCli.Text.Trim();
			if (stNombre.Length == 0) return;
			string stSelect = string.Format("Exec ClienteListaCuenta '{0}', {1}, 0", stNombre, stFiltro);
			int iCuenta = Funcion.iEscalarSQL(cdsGuiaPunto, stSelect, true);
			int IdCliente = 0;
			if (iCuenta >= 0)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(stNombre, 1, 0, 1, 0, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						IdCliente = miBusqueda.idCliente;
						drGuia["NombreCliente"] = miBusqueda.Nombre;
								
						stSelect = "Select IsNull(RUC, '') From Cliente Where idCliente = " + IdCliente.ToString();
						drGuia["rucCliente"] = Funcion.sEscalarSQL(cdsGuiaPunto , stSelect, true);
					}
					else e.Cancel=true;
				}
			}
			else // envia en negativo el idCliente cuando existe solo 1
			{
				IdCliente = -iCuenta;
				stSelect = "Select Nombre From Cliente Where idCliente = " + IdCliente.ToString();
				drGuia["NombreCliente"] = Funcion.sEscalarSQL(cdsGuiaPunto, stSelect, true);

				stSelect = "Select IsNull(RUC, '') From Cliente Where idCliente = " + IdCliente.ToString();
				drGuia["rucCliente"] = Funcion.sEscalarSQL(cdsGuiaPunto , stSelect, true);
			}
			drGuia["idCliente"] = IdCliente;

		}

		private void cdsGuiaPunto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGuiaPunto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProyecto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cmbCarro3_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				cmbCarro3.Value = 0;
				cmbCarro3.Select();
				cmbCarro3.Refresh();
			}
		}

		private void cmbCarro2_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				cmbCarro2.Value = 0;
				cmbCarro2.Select();
				cmbCarro2.Refresh();
			}
		}

		private void cmbCarro1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				cmbCarro1.Value = 0;
				cmbCarro1.Select();
				cmbCarro1.Refresh();
			}
		}

		private void Observacion()
		{
			if (checkContratado.Checked==true) 
			{
				string stSelect= "";
				if (txtTextoTonelaje.Text.Trim().Length>0)
					stSelect=stSelect +  txtTextoTonelaje.Value.ToString().Trim() + " ";   

				if (txtTonPedido.Text.Trim().Length>0)
					stSelect=stSelect + " " + txtTonPedido.Value.ToString().Trim()+ " TN. ";

				if( txtPlaca.Text.Trim().Length >0)
					stSelect=stSelect + " PLACA: " + txtPlaca.Text.ToString()  + " ";   

				drGuia["observaciones"] = stSelect;
			}
			else
			{
				string stSelect= cmbCarro1.Text.Trim();
				if (cmbCarro2.Text.Trim().Length>0)
					stSelect += " con " + cmbCarro2.Text.Trim();
				if (cmbCarro3.Text.Trim().Length>0)
					stSelect += " y " + cmbCarro3.Text.Trim();

				if (txtTextoTonelaje.Text.Trim().Length>0)
					stSelect=stSelect + " " + txtTextoTonelaje.Value.ToString().Trim() + "  ";   

				if (txtTonPedido.Text.Trim().Length>0)
					stSelect=stSelect+ txtTonPedido.Value.ToString().Trim()+ " TN. ";
			
				drGuia["observaciones"] = stSelect;
			}
		}

		private void cmbCarro1_Validated(object sender, System.EventArgs e)
		{
			this.Observacion();
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea copiar este registro?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			int id = this.CopiarAsiento();
			BuscaGuia(id);
		}
		public void BuscaGuia(int idGuia)
		{
			if (idGuia > 0) 
			{
				this.barraDato1.IdRegistro = idGuia;
				this.barraDato1.ProximoId(5);
			}
		}

		private int CopiarAsiento()
		{
			C1.Data.C1DataRow drC = this.cdsGuiaRemisionG.TableViews["GuiaRemisionG"].AddNew();
			drC["contratado"] = drGuia["contratado"];
			drC["finiciotraslado"] = drGuia["finiciotraslado"];
			drC["fTermTraslado"] = drGuia["fTermTraslado"];
			drC["fEmision"] = drGuia["fEmision"];
			drC["fEmision2"] = drGuia["fEmision2"];
			drC["venta"] = drGuia["venta"];
			drC["compra"] = drGuia["compra"];
			drC["transformacion"] = drGuia["transformacion"];
			drC["consignacion"] = drGuia["consignacion"];
			drC["traslado"] = drGuia["traslado"];
			drC["devolucion"] = drGuia["devolucion"];
			drC["importacion"] = drGuia["importacion"];
			drC["exportacion"] = drGuia["exportacion"];
			drC["otros"] = drGuia["otros"];
			drC["idpuntopartida"] = drGuia["idpuntopartida"];
			drC["idpuntollegada"] = drGuia["idpuntollegada"];
			drC["nombrecliente"] = drGuia["nombrecliente"];
	
			drC["diesel"] = drGuia["diesel"];
			drC["adelanto"] = drGuia["adelanto"];
			drC["idcarro1"] = drGuia["idcarro1"];
			drC["idcarro2"] = drGuia["idcarro2"];
			drC["idcarro3"] = drGuia["idcarro3"];
			drC["observaciones"] = drGuia["observaciones"];
			this.cdsGuiaRemisionG.Update();

			// Copia de lineas
			foreach(C1.Data.C1DataRow dr in this.drGuia.GetChildRows())
			{
				C1.Data.C1DataRow drDC = this.cdsGuiaRemisionG.TableViews["GuiaRemisionDetalle"].AddNew();
				drDC["idGuiaRemisionG"] = drC["idGuiaRemisionG"];
				drDC["marca"] = dr["marca"];
				drDC["peso"] = dr["peso"];
				drDC["volumen"] = dr["volumen"];
				drDC["contenido"] = dr["contenido"];
//				drDC["CentroCosto"] = dr["CentroCosto"];
			}
			this.cdsGuiaRemisionG.Update();
			return (int) drC["idGuiaRemisionG"];
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (checkContratado.Checked)
			{
				if (txtCelular.Text.Trim().Length == 0)
				{
				this.errorProvider.SetError(this.txtCelular, "Ingrese celular");
				bOk = false;
			}
				if (txtPlaca.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.txtPlaca, "Ingrese Placa");
					bOk = false;
				}
				if (txtRucChofer.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.txtRucChofer, "Ingrese  Ruc del Chofer");
					bOk = false;
				}
				#region Guias Natranscom
				if (Funcion.bEjecutaSQL(cdsPersonal, "Exec SeteoGExiste 'GNAT'")) 				
				{
					if (cmbChofer.Text.Trim().Length == 0)
					{
						this.errorProvider.SetError(this.cmbChofer, "Escoja Chofer");
						bOk = false;
					}
				}
				#endregion
				else 	if (txtChofer.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.txtChofer, "Ingrese Chofer");
					bOk = false;
				}
			}
			
			if (this.dtFinicioT.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.dtFinicioT, "Ingrese fecha de inicio");
				bOk = false;
			}
//			if (this.dtFTermT.Text.Trim().Length == 0)
//			{
//				this.errorProvider.SetError(this.dtFTermT, "Ingrese fecha de terminacion");
//				bOk = false;
//			}
			if (this.dtFEmision.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.dtFEmision, "Ingrese fecha de emision 1");
				bOk = false;
			}
			if (this.txtCodigo.Text.Trim().Length == 0)
			{
				this.errorProvider.SetError(this.txtCodigo, "Ingrese número de guía de remisión");
				bOk = false;
			}
			else
			{
				int ct=1;
				if (editar==1 && int.Parse(txtCodigo.Text.ToString())==nroguia)
					ct=2;

				string stSelect = "select count(*) from guiaremisiong where nroguia=" + txtCodigo.Text.ToString();   
				int cant=Funcion.iEscalarSQL(cdsGuiaPunto,stSelect);
				if (cant>=ct)
				{
					this.errorProvider.SetError(this.txtCodigo, "Numero de guía ya ha sido utilizado");
					bOk = false;
				}
			}

			if (checkCompra.Checked==false && chkVenta.Checked==false && checkConsignacion.Checked==false && checkDevolucion.Checked==false && checkExportacion.Checked==false && checkImportacion.Checked==false && checkOtros.Checked==false  && checkTransformacion.Checked==false  && checkTraslado.Checked==false)
			{
				this.errorProvider.SetError(this.checkOtros, "Debe escoger algun item");
				bOk = false;
			}

			if (cmbPuntoPartida.Text.Trim().Length==0)
			{
				this.errorProvider.SetError(this.cmbPuntoPartida, "Debe escoger punto partida");
				bOk = false;
			}
			if (cmbPuntoLlegada.Text.Trim().Length==0)
			{
				this.errorProvider.SetError(this.cmbPuntoLlegada, "Debe escoger punto llegada");
				bOk = false;
			}
			if (txtTonPedido.Text.Trim().Length==0)
			{
				this.errorProvider.SetError(this.txtTonPedido, "Debe escribir tonelaje pedido");
				bOk = false;
			}
			if (cmbCarro1.Text.Trim().Length==0)
			{
				this.errorProvider.SetError(this.cmbCarro1, "Debe escoger la maquina1");
				bOk = false;
			}

			return bOk;
		}
		private void BorraErrores()
		{
			this.errorProvider.SetError(this.dtFinicioT, "");
			this.errorProvider.SetError(this.dtFTermT, "");
			this.errorProvider.SetError(this.dtFEmision, "");
			this.errorProvider.SetError(this.txtCodigo, "");
			this.errorProvider.SetError(this.checkOtros, "");
			this.errorProvider.SetError(this.cmbPuntoPartida, "");
			this.errorProvider.SetError(this.cmbPuntoLlegada, "");
			this.errorProvider.SetError(this.cmbCarro1, "");
			this.errorProvider.SetError(this.txtTonPedido, "");
			this.errorProvider.SetError(this.txtCelular, "");
		this.errorProvider.SetError(this.txtPlaca, "");
			this.errorProvider.SetError(this.txtRucChofer, "");
			this.errorProvider.SetError(this.txtChofer, "");
		}

		private void checkContratado_CheckedChanged(object sender, System.EventArgs e)
		{
			if (checkContratado.Checked)
			{
				cmbCarro2.Enabled = false;
				cmbCarro3.Enabled = false;
				txtChofer.Enabled = true;
				txtRucChofer.Enabled = true;
				txtPlaca.Enabled = true;
				txtCelular.Enabled = true;
			}
			else
			{
				cmbCarro2.Enabled = true;
				cmbCarro3.Enabled = true;
				txtChofer.Enabled = false;
				txtRucChofer.Enabled = false;
				txtPlaca.Enabled = false;
				txtCelular.Enabled = false;
			}
		}

		private void checkContratado_Validated(object sender, System.EventArgs e)
		{
			activar();
		}
		private void ChoferCarga()
		{
			
				string stExec = "Exec ChoferCargaNombre " ;
				cmbChofer.DataSource= Funcion.dvProcedimiento(cdsCliente, stExec);
			cmbChofer.DisplayMember = "Nombre";
			cmbChofer.ValueMember = "idCliente";
			//cmbChofer.Rows.Band.Columns["idCliente"].Hidden=true;
			
		}
		private void activar()
		{
			#region Guias Natranscom
			if (Funcion.bEjecutaSQL(cdsPersonal, "Exec SeteoGExiste 'GNAT'")) 				
			{
				
				if ( checkContratado.Checked==true)
				{
				  cmbChofer.Enabled = true;
					ChoferCarga();
					//txtChofer.Visible =false;
					cmbPersonal.Enabled=false;
					txtRucChofer.Enabled=true;  
					txtAdelanto.Enabled=true;
					txtCelular.Enabled=true;
					txtDiesel.Enabled=true;
					txtPlaca.Enabled=true;
					txtComAut.Enabled=false;
					txtComCons.Enabled=false;  
				}
				else
				{
		    	cmbChofer.Enabled = false;
					//txtChofer.Visible=false;
					cmbPersonal.Enabled=true;
					txtRucChofer.Enabled=false;  
					txtAdelanto.Enabled=false;
					txtCelular.Enabled=false;
					txtDiesel.Enabled=false;
					txtPlaca.Enabled=false;
				
					
					txtComAut.Enabled=true;
					txtComCons.Enabled=true;
				}
			}
			#endregion
			if ( checkContratado.Checked==true)
			{
				
				txtChofer.Enabled =true;
				txtRucChofer.Enabled=true;  
				txtAdelanto.Enabled=true;
				txtCelular.Enabled=true;
				txtDiesel.Enabled=true;
				txtPlaca.Enabled=true;
			//	cmbProveedor.Enabled=true;  			

				cmbPersonal.Enabled=false;
				txtComAut.Enabled=false;
				txtComCons.Enabled=false;  
			}
			else
			{
		
				
				txtChofer.Enabled =false;
				txtRucChofer.Enabled=false;  
				txtAdelanto.Enabled=false;
				txtCelular.Enabled=false;
				txtDiesel.Enabled=false;
				txtPlaca.Enabled=false;
				
				cmbPersonal.Enabled=true;
				txtComAut.Enabled=true;
				txtComCons.Enabled=true;
			}
		}

		private void cmbCarro2_Validated(object sender, System.EventArgs e)
		{
				this.Observacion();
		}

		private void cmbCarro3_Validated(object sender, System.EventArgs e)
		{
				this.Observacion();
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			string strFiltro = "{GuiaRemisionG.idGuiaRemisionG} = " 
				+ this.txtIdGuiaRemisionG.Value.ToString();

  		Cursor = Cursors.WaitCursor;
			Reporte miReporte = new Reporte("GuiaRemisionG.rpt",strFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbProveedor_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (BuscaGuiaGruasG miBusqueda = new BuscaGuiaGruasG())
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					if (miBusqueda.idBusca.Text.Length == 0) return;
					int id = int.Parse(miBusqueda.idBusca.Text);
					barraDato1.IdRegistro = id;
					barraDato1.ProximoId(5);
				}
			}

		}

			
		private void btCliente_Click(object sender, System.EventArgs e)
		{
			bool bCliente = !Funcion.EsProveedor(4);
			int IdCliente = 0;
		Cliente miCliente;
			
				miCliente = new Cliente(bCliente, IdCliente);
			
			miCliente.ShowDialog();

				DateTime dtIni = DateTime.Now;
				string stExec = "Exec ChoferCargaNombre ";
				cmbChofer.DataSource = Funcion.dvProcedimiento(cdsCliente, stExec);
				cmbChofer.Refresh();
		
			if (barraDato1.bNuevo && IdCliente == 0)
			{
				try
				{
				//	cmbChofer.EditValue = int.Parse(miCliente.txtChofer.Value.ToString());
				}
				catch{};
			}
			miCliente.Dispose();
			miCliente = null;
			cmbChofer.Select();
			cmbChofer.Refresh();
			
		}

		private void cmbChofer_ValueChanged(object sender, System.EventArgs e)
		{
			string stSelect1;
			string stFiltro = "1";
			
			string stNombre = cmbChofer.Text.Trim();
			if (stNombre.Length == 0) return;
			int IdChofer=(int)cmbChofer.Value;
			string stSelect = string.Format("Exec ClienteListaCuenta '{0}', {1}, 0", IdChofer, stFiltro);
			int iCuenta = Funcion.iEscalarSQL(cdsGuiaPunto, stSelect, true);
			if (iCuenta >= 0)
			{
		 stSelect = "Select IsNull(RUC, '') From Cliente Where idCliente = " + IdChofer;
			drGuia["rucTransp"] = Funcion.sEscalarSQL(cdsCliente, stSelect, true);
				stSelect = "Select IsNull(Nombre, '') From Cliente Where idCliente = " + IdChofer;
				drGuia["NombreTrasnp"] = Funcion.sEscalarSQL(cdsCliente, stSelect, true);
		 stSelect1 = "Select IsNull(Celular, '') From Cliente Where idCliente = " + IdChofer;
			drGuia["Celular"] = Funcion.sEscalarSQL(cdsCliente, stSelect1, true);
		 stSelect = "Select IsNull(Comercial, '') From Cliente Where idCliente = " + IdChofer;
		 drGuia["Placa"] = Funcion.sEscalarSQL(cdsCliente, stSelect, true);
		}
			}

		private void btDigitalizacion_Click(object sender, System.EventArgs e)
		{
		
			int tipo=26;
//			using (Compra miImagen = new Compra(tipo))
//			{
//				miImagen.ShowDialog();
//			}
		}

		private void txtNombreCliD_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stFiltro = "0";
			string stNombre = txtNombreCliD.Text.Trim();
			if (stNombre.Length == 0) return;
			string stSelect = string.Format("Exec ClienteListaCuenta '{0}', {1}, 0", stNombre, stFiltro);
			int iCuenta = Funcion.iEscalarSQL(cdsGuiaPunto, stSelect, true);
			int IdCliente = 0;
			if (iCuenta >= 0)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(stNombre, 1, 0, 0, 0, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						IdCliente = miBusqueda.idCliente;
						drGuia["NombreClienteD"] = miBusqueda.Nombre;
								
						stSelect = "Select IsNull(RUC, '') From Cliente Where idCliente = " + IdCliente.ToString();
						drGuia["rucClienteD"] = Funcion.sEscalarSQL(cdsGuiaPunto , stSelect, true);
					}
					else e.Cancel=true;
				}
			}
			else // envia en negativo el idCliente cuando existe solo 1
			{
				IdCliente = -iCuenta;
				stSelect = "Select Nombre From Cliente Where idCliente = " + IdCliente.ToString();
				drGuia["NombreClienteD"] = Funcion.sEscalarSQL(cdsGuiaPunto, stSelect, true);

				stSelect = "Select IsNull(RUC, '') From Cliente Where idCliente = " + IdCliente.ToString();
				drGuia["rucClienteD"] = Funcion.sEscalarSQL(cdsGuiaPunto , stSelect, true);
			}
			drGuia["idClienteD"] = IdCliente;
		}

	

	}
}
