using System;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de Form1.
	/// </summary>
	public class MenuLatinium : DevExpress.XtraEditors.XtraForm
	{
		#region mnulat
		private System.Windows.Forms.ImageList imageList1;
		private DevExpress.XtraBars.BarButtonItem brAbrir;
		private DevExpress.XtraBars.BarButtonItem brNuevo;
		private DevExpress.XtraBars.BarButtonItem barButtonItem1;
		private DevExpress.XtraBars.BarButtonItem barButtonItem2;
		private DevExpress.XtraBars.BarDockControl barDockControlTop;
		private DevExpress.XtraBars.BarDockControl barDockControlBottom;
		private DevExpress.XtraBars.BarDockControl barDockControlLeft;
		private DevExpress.XtraBars.BarDockControl barDockControlRight;
		private DevExpress.XtraBars.Bar bar1;
		private DevExpress.XtraBars.Bar bar2;
		private DevExpress.XtraBars.BarButtonItem MnuArchNuevo;
		private DevExpress.XtraBars.BarButtonItem MnuArchAbrir;
		private DevExpress.XtraBars.BarSubItem MnuArch;
		private DevExpress.XtraBars.BarButtonItem MnuArchGuardar;
		private DevExpress.XtraBars.BarButtonItem MnuArchPropFact;
		private DevExpress.XtraBars.BarSubItem MnuHerr_;
		private DevExpress.XtraBars.BarButtonItem MnuHerr_Camb_Contr;
		private DevExpress.XtraBars.BarButtonItem MnuHerr_Actu_Vers;
		private DevExpress.XtraBars.BarButtonItem MnuHerr_Cre_Sal_Ini;
		private DevExpress.XtraBars.BarButtonItem MnuHerr_Cop_Sal_Ini;
		private DevExpress.XtraBars.BarButtonItem MnuHerr_Sal_Ini;
		private DevExpress.XtraBars.BarButtonItem MnuHerr_Camb_Ctes;
		private DevExpress.XtraBars.BarButtonItem MnuHerr_Copi_Arch;
		private DevExpress.XtraBars.BarButtonItem MnuHerr_Auditoría;
		private DevExpress.XtraBars.BarButtonItem MnuCont_Ind_Financieros;
		private DevExpress.XtraBars.BarButtonItem MnuFactLot_Ctas_Cob;
		private DevExpress.XtraBars.BarButtonItem MnuFactLot_Ctas_Pag;
		private DevExpress.XtraBars.BarButtonItem MnuFactLot_Facturacion;
		private DevExpress.XtraBars.BarButtonItem MnuFactLot_Transf_Art;
		private DevExpress.XtraBars.BarSubItem MnuVentana;
		private DevExpress.XtraBars.BarSubItem MnuAyuda;
		private DevExpress.XtraBars.BarButtonItem MnuVentana_Casc;
		private DevExpress.XtraBars.BarButtonItem MnuVentana_Org;
		private DevExpress.XtraBars.BarButtonItem MnuVentana_Cicl;
		private DevExpress.XtraBars.BarButtonItem MnuAyuda_Ayuda;
		private DevExpress.XtraBars.BarButtonItem MnuAyudaRegist;
		private DevExpress.XtraBars.BarButtonItem MnuAyudaAcercaCIAD;
		private DevExpress.XtraBars.BarSubItem MnuArchImpoExpo;
		private DevExpress.XtraBars.BarButtonItem MnuArchImpoOtros;
		private DevExpress.XtraBars.BarSubItem MnuArchProp;
		private DevExpress.XtraBars.BarButtonItem MnuArchPropCont;
		private DevExpress.XtraBars.BarButtonItem MnuArchPropRol;
		private DevExpress.XtraBars.BarButtonItem MnuArchConfImp;
		private DevExpress.XtraBars.BarButtonItem MnuArchSalir;
		private DevExpress.XtraBars.BarSubItem MnuEdic;
		private DevExpress.XtraBars.BarButtonItem MnuEdicDesh;
		private DevExpress.XtraBars.BarButtonItem MnuEdicReha;
		private DevExpress.XtraBars.BarButtonItem MnuEdicCortar;
		private DevExpress.XtraBars.BarButtonItem MnuEdicCopi;
		private DevExpress.XtraBars.BarButtonItem MnuEdicPega;
		private DevExpress.XtraBars.BarSubItem MnuCont;
		private DevExpress.XtraBars.BarButtonItem MnuContPlanCtas;
		private DevExpress.XtraBars.BarButtonItem MnuContAsientos;
		private DevExpress.XtraBars.BarButtonItem MnuContRepCont;
		private DevExpress.XtraBars.BarButtonItem MnuContEstadoCta;
		private DevExpress.XtraBars.BarButtonItem MnuContPresupuesto;
		private DevExpress.XtraBars.BarButtonItem MnuContCtas;
		private DevExpress.XtraBars.BarButtonItem MnuContBcos;
		private DevExpress.XtraBars.BarButtonItem MnuContBloqAsientos;
		private DevExpress.XtraBars.BarButtonItem MnuContCierPeriodo;
		private DevExpress.XtraBars.BarButtonItem MnuContAsigCtas;
		private DevExpress.XtraBars.BarButtonItem MnuContFormulario;
		private DevExpress.XtraBars.BarButtonItem MnuFactCompras;
		private DevExpress.XtraBars.BarButtonItem MnuFactOrdComp;
		private DevExpress.XtraBars.BarButtonItem MnuFactVentas;
		private DevExpress.XtraBars.BarButtonItem MnuFactCotizacion;
		private DevExpress.XtraBars.BarButtonItem MnuFactDevVentas;
		private DevExpress.XtraBars.BarButtonItem MnuFactDevCompras;
		private DevExpress.XtraBars.BarButtonItem MnuFactEnvTransferencia;
		private DevExpress.XtraBars.BarButtonItem MnuFactRecTransferencia;
		private DevExpress.XtraBars.BarButtonItem MnuFactReservClientes;
		private DevExpress.XtraBars.BarButtonItem MnuFactReservProveedores;
		private DevExpress.XtraBars.BarButtonItem MnuFactPedEmpleados;
		private DevExpress.XtraBars.BarButtonItem MnuFactPedClientes;
		private DevExpress.XtraBars.BarButtonItem MnuFactPreciosProveedores;
		private DevExpress.XtraBars.BarSubItem MnuFactLotes;
		private DevExpress.XtraBars.BarButtonItem MnuFactArtic;
		private DevExpress.XtraBars.BarButtonItem MnuFactClientes;
		private DevExpress.XtraBars.BarButtonItem MnuFactProvee;
		private DevExpress.XtraBars.BarButtonItem MnuFactGrupArticulos;
		private DevExpress.XtraBars.BarButtonItem MnuFactActFijos;
		private DevExpress.XtraBars.BarButtonItem MnuFactProyectos;
		private DevExpress.XtraBars.BarButtonItem MnuFactCtasCobrar;
		private DevExpress.XtraBars.BarButtonItem MnuFactCtasPagar;
		private DevExpress.XtraBars.BarButtonItem MnuFactRepResumenes;
		private DevExpress.XtraBars.BarButtonItem MnuFactRepInventario;
		private DevExpress.XtraBars.BarSubItem MnuRol;
		private DevExpress.XtraBars.BarButtonItem MnuRolPersonal;
		private DevExpress.XtraBars.BarButtonItem MnuRolCargos;
		private DevExpress.XtraBars.BarButtonItem MnuRolDpto;
		private DevExpress.XtraBars.BarButtonItem MnuRolGenRol;
		private DevExpress.XtraBars.BarButtonItem MnuRolDatGen;
		private DevExpress.XtraBars.BarButtonItem MnuRolRepRol;
		private DevExpress.XtraBars.BarButtonItem MnuRolIncSuel;
		private DevExpress.XtraBars.BarButtonItem MnuRolVerif;
		private DevExpress.XtraBars.BarButtonItem MnuRolBloqRol;
		private DevExpress.XtraBars.BarButtonItem MnuRolAsigCtas;
		private DevExpress.XtraBars.BarSubItem MnuProduc;
		private DevExpress.XtraBars.BarButtonItem MnuProducRecetas;
		private DevExpress.XtraBars.BarButtonItem MnuProducOrdProd;
		private DevExpress.XtraBars.BarButtonItem MnuProducProdTerm;
		private DevExpress.XtraBars.BarButtonItem MnuProducConsReal;
		private DevExpress.XtraBars.BarButtonItem MnuProducActualiza;
		private DevExpress.XtraBars.BarButtonItem MnuProducPrecios;
		private DevExpress.XtraBars.BarButtonItem MnuProducReportes;
		private C1.Data.C1DataSet cdsSeguridad;
		private DevExpress.XtraBars.BarButtonItem MnuProdClientes;
		private DevExpress.XtraBars.BarButtonItem mnuFactTransf;
		private DevExpress.XtraBars.BarButtonItem mnuArchPropEmpresa;
		private DevExpress.XtraBars.BarButtonItem mnuArchPropGen;
		private DevExpress.XtraBars.BarButtonItem barNuevo;
		private DevExpress.XtraBars.BarButtonItem barAbrir;
		private DevExpress.XtraBars.BarButtonItem barGrabar;
		private DevExpress.XtraBars.BarButtonItem MnuEdicSegMenu;
		private DevExpress.XtraBars.BarButtonItem MnuEdicSegUsuario;
		private DevExpress.XtraBars.BarButtonItem MnuArchImpoLat;
		private DevExpress.XtraBars.BarButtonItem MnuFactRepArticulo;
		private DevExpress.XtraBars.BarButtonItem MnuHerr_Act_Saldo;
		private DevExpress.XtraBars.BarButtonItem mnuArchPropProd;
		private DevExpress.XtraBars.BarSubItem MnuContTransacciones;
		private DevExpress.XtraBars.BarButtonItem MnuContTransE;
		private DevExpress.XtraBars.BarButtonItem MnuContTransI;
		private DevExpress.XtraBars.BarButtonItem barAsiento;
		private DevExpress.XtraBars.BarButtonItem BarRepContabilidad;
		private DevExpress.XtraBars.BarButtonItem BarCompras;
		private DevExpress.XtraBars.BarButtonItem BarVentas;
		private DevExpress.XtraBars.BarButtonItem BarRepArticulo;
		private DevExpress.XtraBars.BarButtonItem BarRepResumen;
		private DevExpress.XtraBars.BarButtonItem BarRepInventario;
		private DevExpress.XtraBars.BarButtonItem BarArticulo;
		private DevExpress.XtraBars.BarButtonItem BarCliente;
		private DevExpress.XtraBars.BarButtonItem BarPlanCuentas;
		private DevExpress.XtraBars.BarSubItem MnuFactResPed;
		private DevExpress.XtraBars.BarButtonItem barSeguridad;
		private DevExpress.XtraBars.BarButtonItem MnuContTransR;
		private DevExpress.XtraBars.BarButtonItem mnuHerrGenSQL;
		private DevExpress.XtraBars.BarButtonItem barCobro;
		private DevExpress.XtraBars.BarButtonItem barPagar;
		private DevExpress.XtraBars.BarButtonItem barRolRubro;
		private DevExpress.XtraBars.BarButtonItem mnuHerrReporte;
		private DevExpress.XtraBars.BarButtonItem mnuContCompara;
		private DevExpress.XtraBars.BarButtonItem mnuContAsientoPeriodo;
		private DevExpress.XtraBars.BarButtonItem MnuFactLoteAnalisis;
		private DevExpress.XtraBars.BarButtonItem mnuContLoteDepura;
		private DevExpress.XtraBars.BarButtonItem mnuFactLotOrdenMov;
		private DevExpress.XtraBars.BarButtonItem mnuFactLotAgencia;
		private DevExpress.XtraBars.BarSubItem mnuContLote;
		private System.Windows.Forms.HelpProvider helpProvider1;
		private DevExpress.XtraBars.BarButtonItem MnuFactNotaCV;
		private DevExpress.XtraBars.BarButtonItem MnuFactNotaCC;
		private DevExpress.XtraBars.BarButtonItem MnuVentTrans;
		private DevExpress.XtraBars.BarButtonItem mnuVentOpac;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpOmniLife;
		private DevExpress.XtraBars.BarButtonItem mnuContTransAnexoTr;
		private DevExpress.XtraBars.BarButtonItem mnuFactLotesPago;
		private DevExpress.XtraBars.BarButtonItem MnuFactResPedProv;
		private DevExpress.XtraBars.BarButtonItem mnuContTransFormulario;
		private DevExpress.XtraBars.BarButtonItem mnuFactLotTomaF;
		private DevExpress.XtraBars.BarButtonItem mnuProdOrdenT;
		private DevExpress.XtraBars.BarButtonItem mnuFacLotCxC30;
		private DevExpress.XtraBars.BarButtonItem mnuAyuLiberar;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpTelcosan;
		private DevExpress.XtraBars.BarButtonItem mnuRolesLiquidacion;
		private System.Windows.Forms.Timer timer;
		private DevExpress.XtraBars.BarButtonItem mneArchImpInterActive;
		private DevExpress.XtraBars.BarButtonItem mnuFactLoteTarjeta;
		private DevExpress.XtraBars.BarButtonItem mnuFacLotSerie;
		private DevExpress.XtraBars.BarButtonItem mnuRolAnexo;
		private DevExpress.XtraBars.BarButtonItem MnuContRepGov;
		private DevExpress.XtraBars.BarButtonItem btProdCosto;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpGuiaGruas;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpGruasCompra;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpGuiaVenta;
		private DevExpress.XtraBars.BarButtonItem mnuProdGeneraC;
		private DevExpress.XtraBars.BarButtonItem mnuHerrSegFiltro;
		private DevExpress.XtraBars.BarButtonItem mnuProdAgro;
		private DevExpress.XtraBars.BarButtonItem mnuTarea_Agr;
		private DevExpress.XtraBars.BarButtonItem mnuHerrSegTransf;
		private DevExpress.XtraBars.BarButtonItem mnuHerrSegFiltroSuc;
		private DevExpress.XtraBars.BarButtonItem mbuHerrSegTransf;
		private DevExpress.XtraBars.BarSubItem MnuFact;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpQBI;
		private DevExpress.XtraBars.BarButtonItem mnuHerrSegApReq;
		private DevExpress.XtraBars.BarButtonItem MnuFactResPedRequisicion;
		private DevExpress.XtraBars.BarButtonItem MnuFactResPedComision;
		private DevExpress.XtraBars.BarButtonItem mnuResPedPend;
		private DevExpress.XtraBars.BarButtonItem MnuFactLotimp;
		private DevExpress.XtraBars.BarButtonItem MnuProdSemana;
		private DevExpress.XtraBars.BarButtonItem mnuFactRepComp;
		private DevExpress.XtraBars.BarButtonItem mnuHerrCredito;
		private DevExpress.XtraBars.BarButtonItem mnuHerrBusqueda;
		private DevExpress.XtraBars.BarButtonItem mnuFacLotCierre;
		private DevExpress.XtraBars.BarButtonItem mnuHerrSegReqLis;
		private DevExpress.XtraBars.BarButtonItem mnuHerrSegReqEst;
		private DevExpress.XtraBars.BarSubItem MnuEdicSeguridad;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpCashE;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpCerv;
		private DevExpress.XtraBars.BarButtonItem MnuContFlujo;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpResCDI;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpCostoI;
		private DevExpress.XtraBars.BarButtonItem mnuRolesDistribucion;
		private DevExpress.XtraBars.BarButtonItem mnuFacLotTransfTrans;
		private DevExpress.XtraBars.BarSubItem mnuHerrContratos;
		private DevExpress.XtraBars.BarButtonItem mnuHerrContratosIngreso;
		private DevExpress.XtraBars.BarButtonItem mnuHerrContratosGeneraFact;
		private DevExpress.XtraBars.BarButtonItem mnuHerrContratosAdemdum;
		private DevExpress.XtraBars.BarButtonItem mnuHerrContratosAlarma;
		private DevExpress.XtraBars.BarButtonItem mnuHerrContratosAdemdums;
		private DevExpress.XtraBars.BarButtonItem mnuFacLotCopiaArt;
		private DevExpress.XtraBars.BarButtonItem btControlRemoto;
		private DevExpress.XtraBars.BarButtonItem mnuFacLotFactReembol;
		private DevExpress.XtraBars.BarButtonItem mnuRolesAdicional;
		private DevExpress.XtraBars.BarButtonItem MnuFactResCruzarFactReserva;
		private DevExpress.XtraBars.BarButtonItem mnuRolesAdicionalFormulas;
		private DevExpress.XtraBars.BarButtonItem MnuFactResOTrabajo;
		private DevExpress.XtraBars.BarButtonItem MnuFactResConclusion;
		private DevExpress.XtraBars.BarButtonItem MnuFactResOTrabajoServicio;
		private DevExpress.XtraBars.BarSubItem mnuHerrPresupC;
		private DevExpress.XtraBars.BarButtonItem mnuHerrPresupCIngreso;
		private DevExpress.XtraBars.BarButtonItem mnuHerrPresupCAnalisis;
		private DevExpress.XtraBars.BarButtonItem MnuFactLotimp2;
		private DevExpress.XtraBars.BarButtonItem mnuContTransConciliacion;
		private DevExpress.XtraBars.BarButtonItem mnuHerrPresupCAvanceObra;
		private DevExpress.XtraBars.BarButtonItem mnuHerrPresupCImportar;
		private DevExpress.XtraBars.BarButtonItem MnuFactResJarComision;
		private DevExpress.XtraBars.BarManager lblTiempos;
		private DevExpress.XtraBars.BarEditItem lblMensaje;
		private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit repositoryItemTextEdit1;
		private DevExpress.LookAndFeel.DefaultLookAndFeel skinPrincipal;
		private DevExpress.XtraEditors.Repository.RepositoryItemComboBox repositoryItemComboBox1;
		private DevExpress.XtraBars.BarEditItem barSkin;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpSeprofin;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpMarceloVentas;
		private DevExpress.XtraBars.BarButtonItem mneArchImpYellow;
		private DevExpress.XtraBars.BarButtonItem mnuArchPropConfig;
		private DevExpress.XtraBars.BarButtonItem mnuHerrTester;
		private DevExpress.XtraBars.BarButtonItem mnuContLotCuentaObligatoria;
		private DevExpress.XtraBars.BarButtonItem mnuHerrSegGenGrupos;
		private DevExpress.XtraBars.BarStaticItem mnuArchPropEmpresaR;
		private DevExpress.XtraBars.BarButtonItem MnuEdicSegConfigUsuarios;
		private DevExpress.XtraBars.BarButtonItem MnuFactContometros;
		private DevExpress.XtraBars.BarButtonItem barButtonItem3;
		private DevExpress.XtraBars.BarSubItem barSubItem1;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpJuanMarcet;
		private DevExpress.XtraBars.BarButtonItem barButtonItem4;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpFactRet;
		private DevExpress.XtraBars.BarButtonItem mnuHerrAgenda;
		private DevExpress.XtraBars.BarButtonItem mnuHerrPresupCSolicitudPago;
		private DevExpress.XtraBars.BarButtonItem mnuHerrPresupCReportes;
		private DevExpress.XtraBars.BarButtonItem mnuAyuJuego;
		private DevExpress.XtraBars.BarButtonItem mnuHerrContratosAmortiza;
		private DevExpress.XtraBars.BarButtonItem MnuFactPreFactura;
		private DevExpress.XtraBars.BarButtonItem MnuDepTec;
		private DevExpress.XtraBars.BarButtonItem barButtonItem5;
		private DevExpress.XtraBars.BarSubItem barSubItem2;
		private DevExpress.XtraBars.BarButtonItem MnuHerrAlarma;
		private DevExpress.XtraBars.BarButtonItem MnuFacSpif;
		private DevExpress.XtraBars.BarButtonItem MnuFacSerie;
		private DevExpress.XtraBars.BarButtonItem MnnuFactLotesPromociones;
		private DevExpress.XtraBars.BarButtonItem barButtonItem6;
		private DevExpress.XtraBars.BarSubItem MnuCredito;
		private DevExpress.XtraBars.BarButtonItem MnuCreditoSolRapida;
		private DevExpress.XtraBars.BarButtonItem MnuFactLoteTransMult;
		private DevExpress.XtraBars.BarButtonItem MnuFactLoteSpif;
		private DevExpress.XtraBars.BarButtonItem barButtonItem11;
		private DevExpress.XtraBars.BarButtonItem MnuFacInv;
		private DevExpress.XtraBars.BarButtonItem MnuFacOCImp;
		private DevExpress.XtraBars.BarButtonItem MnuCreditoCotizador;
		private DevExpress.XtraBars.BarButtonItem MnuFactLotesPromVen;
		private DevExpress.XtraBars.BarButtonItem MnuCreditoCuota;
		private DevExpress.XtraBars.BarButtonItem MnuCreGastoCobran;
		private DevExpress.XtraBars.BarButtonItem MnuFactCombos;
		private DevExpress.XtraBars.BarButtonItem MnuCreCobranza;
		private DevExpress.XtraBars.BarButtonItem MnuFactAnticipos;
		private DevExpress.XtraBars.BarButtonItem MnuFacComConfiguracion;
		private DevExpress.XtraBars.BarButtonItem MnuFactComCalificacion;
		private DevExpress.XtraBars.BarButtonItem MnuFactDetalle;
		private DevExpress.XtraBars.BarButtonItem MnuFactComConsulta;
		private DevExpress.XtraBars.BarButtonItem MnuFacArtServicios;
		private DevExpress.XtraBars.BarButtonItem MnuFactSerie;
		private DevExpress.XtraBars.BarButtonItem mnuFactBodegas;
		private DevExpress.XtraBars.BarSubItem barSubItem6;
		private DevExpress.XtraBars.BarSubItem barSubItem7;
		private DevExpress.XtraBars.BarButtonItem MnuFactConsignacion;
		private DevExpress.XtraBars.BarButtonItem MnuFacPagosAut;
		private DevExpress.XtraBars.BarButtonItem MnuFactAutPagos;
		private DevExpress.XtraBars.BarButtonItem MnuContAsigArt;
		private DevExpress.XtraBars.BarButtonItem MnuFacSeriales;
		private DevExpress.XtraBars.BarButtonItem MnuContMantTarj;
		private DevExpress.XtraBars.BarButtonItem MnuFactEstTransf;
		private DevExpress.XtraBars.BarButtonItem FacGuiaRemision;
		private DevExpress.XtraBars.BarButtonItem mnuFactGuiaRemision;
		private DevExpress.XtraBars.BarButtonItem mnuFactActPrecio;
		private DevExpress.XtraBars.BarButtonItem MnuProdEstado;
		private DevExpress.XtraBars.BarButtonItem MnuContDepositoVouchers;
		private DevExpress.XtraBars.BarButtonItem mnuCreLotesUnibanco;
		private DevExpress.XtraBars.BarButtonItem MnuContCuadreCaja;
		private DevExpress.XtraBars.BarButtonItem MnuContCXP;
		private DevExpress.XtraBars.BarButtonItem MnuContCXC;
		private DevExpress.XtraBars.BarButtonItem mnuFactVentasConsignacion;
		private DevExpress.XtraBars.BarButtonItem MnuDTSTA;
		private DevExpress.XtraBars.BarButtonItem MnuDTSolicitud;
		private DevExpress.XtraBars.BarButtonItem barButtonItem15;
		private DevExpress.XtraBars.BarButtonItem barButtonItem16;
		private DevExpress.XtraBars.BarSubItem MnuBodega;
		private DevExpress.XtraBars.BarSubItem MnuContSRI;
		private DevExpress.XtraBars.BarSubItem MnuCompras;
		private DevExpress.XtraBars.BarSubItem mnuVentasComisiones;
		private DevExpress.XtraBars.BarSubItem mnuVentas;
		private DevExpress.XtraBars.BarButtonItem mnuCrediPointSolicitud;
		private DevExpress.XtraBars.BarButtonItem mnuCrediPointMantenedor;
		private DevExpress.XtraBars.BarButtonItem mnuCrediPointVerTerreno;
		private DevExpress.XtraBars.BarSubItem mnuCrediPointGestiones;
		private DevExpress.XtraBars.BarButtonItem mnuCrediPointGestVerCed;
		private DevExpress.XtraBars.BarButtonItem mnuCrediPointGestRevCed;
		private DevExpress.XtraBars.BarSubItem mnuCrediPointCatalogos;
		private DevExpress.XtraBars.BarButtonItem mnuProduccionPendientes;
		private DevExpress.XtraBars.BarSubItem mnuDT;
		private DevExpress.XtraBars.BarButtonItem mnuVentasResumenDiario;
		private DevExpress.XtraBars.BarButtonItem mnuVentasCB;
		private DevExpress.XtraBars.BarButtonItem mnuVentasCotizacion;
		private DevExpress.XtraBars.BarButtonItem mnuVentaAutDistribuidor;
		private DevExpress.XtraBars.BarButtonItem mnuVentaLiqCons;
		private DevExpress.XtraBars.BarButtonItem mnuComprasComboZOOM;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaKardex;
		private DevExpress.XtraBars.BarButtonItem mnuComprasLiquidaGastos;
		private DevExpress.XtraBars.BarButtonItem mnuMatrizDsctos;
		private DevExpress.XtraBars.BarButtonItem mnuComprasAgenda;
		private DevExpress.XtraBars.BarButtonItem mnuCompraGestion;
		private DevExpress.XtraBars.BarButtonItem mnuVentasComDetallePersonal;
		private DevExpress.XtraBars.BarButtonItem mnuContResCXC;
		private DevExpress.XtraBars.BarButtonItem mnuContResCXP;
		private DevExpress.XtraBars.BarButtonItem mnuContRepAnt;
		private DevExpress.XtraBars.BarButtonItem barButtonItem7;
		private DevExpress.XtraBars.BarButtonItem mnuRolFiniquitos;
		private DevExpress.XtraBars.BarButtonItem barButtonItem8;
		private DevExpress.XtraBars.BarButtonItem mnuHerramientasNumeracion;
		private DevExpress.XtraBars.BarButtonItem mnuContCCxCCP;
		private DevExpress.XtraBars.BarButtonItem barButtonItem10;
		private DevExpress.XtraBars.BarSubItem barSubItem3;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReporteCompara;
		private DevExpress.XtraBars.BarButtonItem mnuContabilidadAsDesc;
		private DevExpress.XtraBars.BarButtonItem mnuContAntProv;
		private DevExpress.XtraBars.BarButtonItem mnuContRepCC;
		private DevExpress.XtraBars.BarButtonItem mnuContASF;
		private DevExpress.XtraBars.BarButtonItem mnuContMovBanc;
		private DevExpress.XtraBars.BarButtonItem mnuContEstadoCheques;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaReportes;
		private DevExpress.XtraBars.BarButtonItem mnuTomaInventario;
		private DevExpress.XtraBars.BarButtonItem mnuComprasRepSpifVend;
		private DevExpress.XtraBars.BarButtonItem mnuContPrestamos;
		private DevExpress.XtraBars.BarButtonItem mnuContTEB;
		private DevExpress.XtraBars.BarButtonItem mnuVentasDA;
		private DevExpress.XtraBars.BarButtonItem barButtonItem13;
		private DevExpress.XtraBars.BarSubItem barSubItem4;
		private DevExpress.XtraBars.BarButtonItem mnuContConsumosTC;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaDCI;
		private DevExpress.XtraBars.BarButtonItem mnuContLCTC;
		private DevExpress.XtraBars.BarButtonItem mnuProduccionTPT;
		private DevExpress.XtraBars.BarButtonItem mnuContTCC;
		private DevExpress.XtraBars.BarButtonItem barButtonItem14;
		private DevExpress.XtraBars.BarButtonItem mnuDTIngresoMultiple;
		private DevExpress.XtraBars.BarButtonItem mnuDTRetiro;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaCI;
		private DevExpress.XtraBars.BarButtonItem mnuVentasImpresionDePrecios;
		private DevExpress.XtraBars.BarButtonItem barButtonItem9;
		private DevExpress.XtraBars.BarButtonItem mnuHerrSegCD;
		private DevExpress.XtraBars.BarButtonItem mnuCmpRepVM;
		private DevExpress.XtraBars.BarButtonItem mnuRepCmpTransf;
		private DevExpress.XtraBars.BarButtonItem barButtonItem12;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaCierreInventario;
		private DevExpress.XtraBars.BarButtonItem barButtonItem17;
		private DevExpress.XtraBars.BarButtonItem barButtonItem18;
		private DevExpress.XtraBars.BarButtonItem mnuGPPD;
		private DevExpress.XtraBars.BarButtonItem barButtonItem19;
		private DevExpress.XtraBars.BarSubItem mnuCPReportes;
		private DevExpress.XtraBars.BarButtonItem mnuCPReportesCliPoint;
		private DevExpress.XtraBars.BarButtonItem mnuCPInformeCobrador;
		private DevExpress.XtraBars.BarButtonItem mnuCompraReportesCPA;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesPorMarca;
		private DevExpress.XtraBars.BarSubItem mnuComprasReportes;
		private DevExpress.XtraBars.BarButtonItem mnuHerramientasCRIVA;
		private DevExpress.XtraBars.BarButtonItem mnuComprasPromociones;
		private DevExpress.XtraBars.BarButtonItem mnuVentaPromocionesActivas;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaIF;
		private DevExpress.XtraBars.BarButtonItem mnuVentasRVML;
		private DevExpress.XtraBars.BarButtonItem mnuCrediPointDFG;
		private DevExpress.XtraBars.BarButtonItem mnuCrediPointMEF;
		private DevExpress.XtraBars.BarButtonItem mnuCPCP;
		private DevExpress.XtraBars.BarButtonItem mnuCPIO;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaSectorizacion;
		private DevExpress.XtraBars.BarButtonItem mnuComprasTEL;
		private System.Windows.Forms.ToolTip toolTip1;
		private DevExpress.XtraBars.BarButtonItem mnuVentasIPP;
		private DevExpress.XtraBars.BarButtonItem mnuContBancosCBA;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaInvVal;
		private DevExpress.XtraBars.BarButtonItem mnuCPLCS;
		private DevExpress.XtraBars.BarButtonItem mnuCPRepEquifax;
		private DevExpress.XtraBars.BarButtonItem mnuCPMFCP;
		private DevExpress.XtraBars.BarButtonItem mnuVentasRepMU;
		private DevExpress.XtraBars.BarButtonItem mnuContRepGastos;
		private DevExpress.XtraBars.BarButtonItem mnuCPRepCP;
		private DevExpress.XtraBars.BarButtonItem barButtonItem20;
		private DevExpress.XtraBars.BarButtonItem mnuContRepFacElec;
		private DevExpress.XtraBars.BarToolbarsListItem barToolbarsListItem1;
		private DevExpress.XtraBars.BarButtonItem barButtonItem21;
		private DevExpress.XtraBars.BarSubItem barSubItem8;
		private DevExpress.XtraBars.BarSubItem mnuLotte;
		private DevExpress.XtraBars.BarButtonItem mnuTmaHabitacion;
		private DevExpress.XtraBars.BarButtonItem mnuTmpEstadia;
		private DevExpress.XtraBars.BarButtonItem mnuTmpHabitaciones;
		private DevExpress.XtraBars.BarButtonItem mnuCuadroHabitaciones;
		private DevExpress.XtraBars.BarButtonItem mnuCobrPrenda;
		private DevExpress.XtraBars.BarButtonItem mnuKardexLot;
		private DevExpress.XtraBars.BarButtonItem mnuPromociones;
		private DevExpress.XtraBars.BarButtonItem mnuReservas;
		private DevExpress.XtraBars.BarButtonItem mnuHabitaciones;
		private DevExpress.XtraBars.BarButtonItem mnuCpLbp;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesAIM;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesRF;
		private DevExpress.XtraBars.BarButtonItem barButtonItem22;
		private DevExpress.XtraBars.BarButtonItem barButtonItem23;
		private DevExpress.XtraBars.BarSubItem barSubItem9;
		private DevExpress.XtraBars.BarButtonItem barButtonItem24;
		private DevExpress.XtraBars.BarButtonItem barButtonItem25;
		private DevExpress.XtraBars.BarButtonItem barButtonItem26;
		private DevExpress.XtraBars.BarButtonItem mnuVentasVentasCP;
		private DevExpress.XtraBars.BarButtonItem barButtonItem27;
		private DevExpress.XtraBars.BarButtonItem mnuHerrALF;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaIngSeriales;
		private DevExpress.XtraBars.BarButtonItem barButtonItem28;
		private DevExpress.XtraBars.BarButtonItem barButtonItem29;
		private DevExpress.XtraBars.BarButtonItem barButtonItem30;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaGRC;
		private DevExpress.XtraBars.BarButtonItem barButtonItem31;
		private DevExpress.XtraBars.BarButtonItem barButtonItem32;
		private DevExpress.XtraBars.BarButtonItem mnuCPRRCS;
		private DevExpress.XtraBars.BarButtonItem mnuComprasGDA;
		private DevExpress.XtraBars.BarButtonItem mnuLoteNotasDeEntrega;
		private DevExpress.XtraBars.BarButtonItem mnuCPRD;
		private DevExpress.XtraBars.BarButtonItem mnuCPRepRDD;
		private DevExpress.XtraBars.BarButtonItem mnuCPVEP;
		private DevExpress.XtraBars.BarButtonItem mnuVentasARP;
		private DevExpress.XtraBars.BarButtonItem mnuCPRepECP;
		private DevExpress.XtraBars.BarButtonItem mnuCPEDC;
		private DevExpress.XtraBars.BarButtonItem mnuComprasRepFP;
		private DevExpress.XtraBars.BarButtonItem barButtonItem33;
		private DevExpress.XtraBars.BarButtonItem barButtonItem34;
		private DevExpress.XtraBars.BarButtonItem barButtonItem36;
		private DevExpress.XtraBars.BarButtonItem mnuCPHRC;
		private DevExpress.XtraBars.BarButtonItem mnuVentaRMU;
		private DevExpress.XtraBars.BarButtonItem mnuCPRRC;
		private DevExpress.XtraBars.BarButtonItem barButtonItem37;
		private DevExpress.XtraBars.BarButtonItem barButtonItem38;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRM;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRMP;
		private DevExpress.XtraBars.BarButtonItem mnuRolesCom_AV;
		private DevExpress.XtraBars.BarButtonItem mnuRolesCom_PA;
		private DevExpress.XtraBars.BarButtonItem mnuRolesCom_PL;
		private DevExpress.XtraBars.BarButtonItem mnuRolesCom_PML;
		private DevExpress.XtraBars.BarSubItem mnuRolesComisiones;
		private DevExpress.XtraBars.BarButtonItem mnuVentasPC;
		private DevExpress.XtraBars.BarButtonItem barButtonItem35;
		private DevExpress.XtraBars.BarButtonItem mnuRolPA;
		private DevExpress.XtraBars.BarButtonItem barButtonItem39;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobranzasRCP;
		private DevExpress.XtraBars.BarButtonItem mnuRolComisionesPC;
		private DevExpress.XtraBars.BarButtonItem barButtonItem40;
		private DevExpress.XtraBars.BarButtonItem mnuVentasResumenComisiones;
		private DevExpress.XtraBars.BarButtonItem mnuCPReportesRCV;
		private DevExpress.XtraBars.BarButtonItem mnuRolesPagoComisiones;
		private DevExpress.XtraBars.BarButtonItem mnuRolesComAJV;
		private DevExpress.XtraBars.BarButtonItem mnuCPDFGP;
		private DevExpress.XtraBars.BarButtonItem mnuComprasAPC;
		private DevExpress.XtraBars.BarButtonItem barButtonItem41;
		private DevExpress.XtraBars.BarSubItem mnuRolesReportes;
		private DevExpress.XtraBars.BarButtonItem mnuRolesReportesPVF;
		private DevExpress.XtraBars.BarButtonItem barButtonItem42;
		private DevExpress.XtraBars.BarButtonItem mnuCom_PAGE;
		private DevExpress.XtraBars.BarButtonItem mnuCom_PMLGE;
		private DevExpress.XtraBars.BarButtonItem mnuCom_PCGE;
		private DevExpress.XtraBars.BarButtonItem mnuVentasRepGE;
		private DevExpress.XtraBars.BarSubItem mnuCom_Reportes;
		private DevExpress.XtraBars.BarButtonItem mnuCom_ReportesRM;
		private DevExpress.XtraBars.BarButtonItem mnuCobranzasRGC;
		private DevExpress.XtraBars.BarButtonItem mnuRolesReportesVXM;
		private DevExpress.XtraBars.BarButtonItem mnuVentasMarcador;
		private DevExpress.XtraBars.BarSubItem mnuCPCobranzas;
		private DevExpress.XtraBars.BarButtonItem barButtonItem44;
		private DevExpress.XtraBars.BarSubItem mnuReportesBodega;
		private DevExpress.XtraBars.BarSubItem mnuContReportes;
		private DevExpress.XtraBars.BarButtonItem barButtonItem45;
		private DevExpress.XtraBars.BarButtonItem barButtonItem46;
		private DevExpress.XtraBars.BarButtonItem barButtonItem47;
		private DevExpress.XtraBars.BarSubItem barSubItem10;
		private DevExpress.XtraBars.BarButtonItem barButtonItem48;
		private DevExpress.XtraBars.BarSubItem barSubItem11;
		private DevExpress.XtraBars.BarButtonItem mnuCPLotesPPagos;
		private DevExpress.XtraBars.BarButtonItem barButtonItem49;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesComparativos;
		private DevExpress.XtraBars.BarButtonItem mnuCPReportesGDCM;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaIPR;
		private DevExpress.XtraBars.BarButtonItem barButtonItem50;
		private DevExpress.XtraBars.BarButtonItem mnuContRepCajasChicas;
		private DevExpress.XtraBars.BarButtonItem barButtonItem55;
		private DevExpress.XtraBars.BarButtonItem mnuCPReportesAbonos;
		private DevExpress.XtraBars.BarButtonItem mnuCPReportesNDC;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesVentasPorMes;
		private DevExpress.XtraBars.BarButtonItem mnuContReportesGA;
		private DevExpress.XtraBars.BarButtonItem barButtonItem56;
		private DevExpress.XtraBars.BarButtonItem mnuContBancosCP;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesTC;
		private DevExpress.XtraBars.BarButtonItem mnuRolesReportesUniformes;
		private DevExpress.XtraBars.BarButtonItem barButtonItem57;
		private DevExpress.XtraBars.BarButtonItem mnuContBancosDPL;
		private DevExpress.XtraBars.BarButtonItem mnuContBancosCDPL;
		private DevExpress.XtraBars.BarButtonItem barButtonItem61;
		private DevExpress.XtraBars.BarButtonItem barButtonItem66;
		private DevExpress.XtraBars.BarButtonItem mnuVentasRepRVF;
		private DevExpress.XtraBars.BarButtonItem mnuVentasPlanifVentas;
		private DevExpress.XtraBars.BarButtonItem barButtonItem67;
		private DevExpress.XtraBars.BarButtonItem mnuDTProductosDeFeria;
		private DevExpress.XtraBars.BarButtonItem barButtonItem69;
		private DevExpress.XtraBars.BarButtonItem mnuVentasPlanifJefeVentas;
		private DevExpress.XtraBars.BarButtonItem barButtonItem68;
		private DevExpress.XtraBars.BarButtonItem mnuRolesReportesNominaGeneral;
		private DevExpress.XtraBars.BarButtonItem barButtonItem70;
		private DevExpress.XtraBars.BarButtonItem barButtonItem71;
		private DevExpress.XtraBars.BarButtonItem mnuRolesCom_EPV;
		private DevExpress.XtraBars.BarButtonItem mnuRolPersonalPrueba;
		private DevExpress.XtraBars.BarButtonItem mnuRolComPIA;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesDTC;
		private DevExpress.XtraBars.BarButtonItem mnuVRE;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesEncuestas;
		private DevExpress.XtraBars.BarButtonItem barButtonItem74;
		private DevExpress.XtraBars.BarSubItem barSubItem14;
		private DevExpress.XtraBars.BarButtonItem mnuLoteRepProcSIM;
		private DevExpress.XtraBars.BarButtonItem mnuCompraCSA;
		private DevExpress.XtraBars.BarButtonItem barButtonItem75;
		private DevExpress.XtraBars.BarButtonItem mnuVentasGCCC;
		private DevExpress.XtraBars.BarButtonItem barButtonItem76;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaRM;
		private DevExpress.XtraBars.BarButtonItem mnuCPReportesBR;
		private DevExpress.XtraBars.BarButtonItem mnuHerramientasPTP;
		private DevExpress.XtraBars.BarButtonItem mnuHerramientasAP;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaRepKS;
		private DevExpress.XtraBars.BarButtonItem barButtonItem72;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRAT;
		private DevExpress.XtraBars.BarButtonItem mnuDTIREP;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRRPT;
		private DevExpress.XtraBars.BarButtonItem mnuDTIREPC;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRPPL;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRAPD;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRSNC;
		private DevExpress.XtraBars.BarSubItem mnuDTIRREP;
		private DevExpress.XtraBars.BarSubItem mnuDTIRRMA;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRFDT;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRSIN;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRCDE;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRRRP;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRRTE;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRRPG;
		private DevExpress.XtraBars.BarButtonItem mnuDTIRRSR;
		private DevExpress.XtraBars.BarButtonItem barButtonItem51;
		private DevExpress.XtraBars.BarButtonItem mnuRolCom_PorcCarpa;
		private DevExpress.XtraBars.BarButtonItem mnuReporteVentasMensualesPorLocal;
		private DevExpress.XtraBars.BarButtonItem mnuVentasIPR;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaCDI;
		private DevExpress.XtraBars.BarButtonItem frmDTTransfLocales;
		private DevExpress.XtraBars.BarButtonItem mnuDTRA;
		private DevExpress.XtraBars.BarButtonItem mnuDTReporteEstadoIngresos;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobranzasLPCA;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesVA;
		private DevExpress.XtraBars.BarButtonItem mnuDTAprobacionRemates;
		private DevExpress.XtraBars.BarButtonItem mnuDTASSF;
		private DevExpress.XtraBars.BarButtonItem barButtonItem58;
		private DevExpress.XtraBars.BarSubItem mnuMetricas;
		private DevExpress.XtraBars.BarButtonItem barButtonItem59;
		private DevExpress.XtraBars.BarSubItem mnuMetricasRRHH;
		private DevExpress.XtraBars.BarButtonItem barButtonItem60;
		private DevExpress.XtraBars.BarButtonItem barButtonItem62;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaReportesECDI;
		private DevExpress.XtraBars.BarButtonItem barButtonItem63;
		private DevExpress.XtraBars.BarButtonItem mnuSTA;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesFFP;
		private DevExpress.XtraBars.BarButtonItem mnuVentasEOC;
		private DevExpress.XtraBars.BarButtonItem mnuLottePlanificacion;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesVPD;
		private DevExpress.XtraBars.BarButtonItem mnuVentasGCR;
		private DevExpress.XtraBars.BarButtonItem mnuLotteReportesIT;
		private DevExpress.XtraBars.BarButtonItem barButtonItem73;
		private DevExpress.XtraBars.BarButtonItem mnuLotteCompetencia;
		private DevExpress.XtraBars.BarButtonItem mnuLotteReporteIPH;
		private DevExpress.XtraBars.BarButtonItem mnuRolReportesAE;
		private DevExpress.XtraBars.BarButtonItem mnuContabilidadFlujoPMP;
		private DevExpress.XtraBars.BarButtonItem mnuCrediPOINTRMR;
		private DevExpress.XtraBars.BarButtonItem mnuCrediPOINTReportesRMR;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesCXC;
		private DevExpress.XtraBars.BarButtonItem mnuVentasPFR;
		private DevExpress.XtraBars.BarButtonItem mnuCPReportesEC;
		private DevExpress.XtraBars.BarButtonItem mnuPSRDC;
		private DevExpress.XtraBars.BarButtonItem mnuRolesReportesHHT;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobranzasCCPPC;
		private DevExpress.XtraBars.BarButtonItem mnuCPReportesREC;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesVPA;
		private DevExpress.XtraBars.BarButtonItem mnuPCGEx;
		private DevExpress.XtraBars.BarButtonItem mnuVentasRP;
		private DevExpress.XtraBars.BarButtonItem mnuVentasECW;
		private DevExpress.XtraBars.BarButtonItem mnuHerrVerifDepositos;
		private DevExpress.XtraBars.BarButtonItem mnuHerrSegClientes;
		private DevExpress.XtraBars.BarButtonItem mnuHerrDepositos;
		private DevExpress.XtraBars.BarButtonItem mnuContabilidadFPPP;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesVentasPorArticulos;
		private DevExpress.XtraBars.BarButtonItem mnuVentasReportesVentasarticuloPorMes;
		private DevExpress.XtraBars.BarButtonItem mnuComprasRepCodigoArticulos;
		private DevExpress.XtraBars.BarButtonItem mnuDTSubGrupo;
		private DevExpress.XtraBars.BarButtonItem barButtonItem52;
		private DevExpress.XtraBars.BarButtonItem mnuComprasCatalogo;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaRepMD;
		private DevExpress.XtraBars.BarButtonItem mnuContFPPP;
		private DevExpress.XtraBars.BarSubItem mnuContabilidadProvisiones;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobranzasEDC;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaTIM;
		private DevExpress.XtraBars.BarButtonItem mnuContRB;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaTRM;
		private DevExpress.XtraBars.BarButtonItem mnuVentasRPA;
		private DevExpress.XtraBars.BarButtonItem mnuRolesComisionesPS;
		private DevExpress.XtraBars.BarButtonItem mnuRolesComisionesPCS;
		private DevExpress.XtraBars.BarButtonItem barButtonItem53;
		private DevExpress.XtraBars.BarButtonItem mnuTarjetasDeIngreso;
		private DevExpress.XtraBars.BarButtonItem mnuVentasCNOS;
		private DevExpress.XtraBars.BarButtonItem mnuDirectorio;
		private DevExpress.XtraBars.BarSubItem mnuCPCartera;
		private DevExpress.XtraBars.BarButtonItem mnuCPReportesRRM;
		private DevExpress.XtraBars.BarButtonItem mnuComprasProformas;
		private DevExpress.XtraBars.BarButtonItem mnuRolReportesIND;
		private DevExpress.XtraBars.BarButtonItem mnuCPReportesCVPM;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobranzasMRPR;
		private DevExpress.XtraBars.BarButtonItem mnuHerramientasPTA;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobranzasGDC;
		private DevExpress.XtraBars.BarButtonItem mnuVentaRMUA;
		private DevExpress.XtraBars.BarButtonItem mnuContRepEstadoCB;
		private DevExpress.XtraBars.BarButtonItem mnuContEstadosCajas;
		private DevExpress.XtraBars.BarButtonItem mnuContIngresoDiarioCaja;
		private DevExpress.XtraBars.BarButtonItem mnuContIngresosDeCajas;
		private DevExpress.XtraBars.BarButtonItem mnuContEstadoGeneralDeCaja;
		private DevExpress.XtraBars.BarSubItem mnuContCaja;
		private DevExpress.XtraBars.BarButtonItem mnuContRepPresupAlm;
		private DevExpress.XtraBars.BarButtonItem mnuHerramientasAPL;
		private DevExpress.XtraBars.BarButtonItem mnuHerramientasASDP;
		private DevExpress.XtraBars.BarButtonItem mnuCPRevision;
		private DevExpress.XtraBars.BarButtonItem barButtonItem43;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobranzasCarteraAIP;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobranzasRPE;
		private DevExpress.XtraBars.BarButtonItem mnuRolesDecimoTercerSueldo;
		private DevExpress.XtraBars.BarButtonItem barButtonItem54;
		private DevExpress.XtraBars.BarSubItem mnuCPCRD;
		private DevExpress.XtraBars.BarButtonItem mnuCPCDRDC;
		private DevExpress.XtraBars.BarButtonItem mnuCPCDRVC;
		private DevExpress.XtraBars.BarButtonItem mnuCPCDSR;
		private DevExpress.XtraBars.BarButtonItem mnuCPCDRSLT;
		private DevExpress.XtraBars.BarButtonItem mnuCPCDLCPG;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaReportesKardexCosto;
		private DevExpress.XtraBars.BarButtonItem barButtonItem64;
		private DevExpress.XtraBars.BarLinkContainerItem barLinkContainerItem1;
		private DevExpress.XtraBars.BarSubItem barSubItem5;
		private DevExpress.XtraBars.BarButtonItem barButtonItem65;
		private DevExpress.XtraBars.BarSubItem mnuRPtBodega;
		private DevExpress.XtraBars.BarSubItem mnuRPtContabilidad;
		private DevExpress.XtraBars.BarButtonItem barButtonItem77;
		private DevExpress.XtraBars.BarButtonItem barButtonItem78;
		private DevExpress.XtraBars.BarSubItem mnuRPtRoles;
		private DevExpress.XtraBars.BarButtonItem barButtonItem79;
		private DevExpress.XtraBars.BarSubItem mnuRPtCompras;
		private DevExpress.XtraBars.BarButtonItem barButtonItem80;
		private DevExpress.XtraBars.BarSubItem mnuRPtVentas;
		private DevExpress.XtraBars.BarSubItem mnuRPtCREPOINT;
		private DevExpress.XtraBars.BarSubItem mnuRPtDT;
		private DevExpress.XtraBars.BarSubItem mnuRPtPeonyServicios;
		private DevExpress.XtraBars.BarButtonItem mnuRPtBodegaInvFEcha;
		private DevExpress.XtraBars.BarButtonItem mnuRPtBodegaRptTrans;
		private DevExpress.XtraBars.BarButtonItem mnuRPtBodegaInvVal;
		private DevExpress.XtraBars.BarButtonItem mnuRPtBodegaKXSer;
		private DevExpress.XtraBars.BarButtonItem mnuRPtBodegaEsConInv;
		private DevExpress.XtraBars.BarButtonItem mnuRPtBodegaMovDesc;
		private DevExpress.XtraBars.BarButtonItem mnuRPtContabilidadCont;
		private DevExpress.XtraBars.BarButtonItem mnuRPtContabilidadRet;
		private DevExpress.XtraBars.BarButtonItem mnuRPtContabilidadAsDEsc;
		private DevExpress.XtraBars.BarButtonItem mnuRPtContabilidadGast;
		private DevExpress.XtraBars.BarButtonItem mnuRPtContabilidadFactEl;
		private DevExpress.XtraBars.BarButtonItem mnuRPtContabilidadCaCh;
		private DevExpress.XtraBars.BarButtonItem mnuRPtContabilidadGaAn;
		private DevExpress.XtraBars.BarButtonItem mnuRPtContabilidadEstaConB;
		private DevExpress.XtraBars.BarButtonItem mnuRPtContabilidadPrXAlm;
		private DevExpress.XtraBars.BarButtonItem mnuRPtRolesPFA;
		private DevExpress.XtraBars.BarButtonItem mnuRPtRolesVeXM;
		private DevExpress.XtraBars.BarButtonItem mnuRPtRolesUni;
		private DevExpress.XtraBars.BarButtonItem mnuRPtRolesNG;
		private DevExpress.XtraBars.BarButtonItem mnuRPtRolesAEM;
		private DevExpress.XtraBars.BarButtonItem mnuRPtRolesHHT;
		private DevExpress.XtraBars.BarButtonItem mnuRPtRolesIXD;
		private DevExpress.XtraBars.BarButtonItem mnuRPtComprasCPA;
		private DevExpress.XtraBars.BarButtonItem mnuRPtComprasT;
		private DevExpress.XtraBars.BarButtonItem mnuRPtComprasRAC;
		private DevExpress.XtraBars.BarButtonItem mnuRPtComprasCA;
		private DevExpress.XtraBars.BarButtonItem mnuIngreso;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasPorcMargUtilidad;
		private DevExpress.XtraBars.BarButtonItem mnuRPtBodegaKCC;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasPorcMarca;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasPorcMarcaUtilid;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasMargUti;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasReMarUti;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasMaUXArt;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasAnInMEn;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasRevFact;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasXDia;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasNFactPForPago;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasGaraExt;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasXFecha;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasComDeVentas;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasANuXLo;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasMenXLo;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasTarCredi;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasDetTc;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasEncuestas;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasReCXCo;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasXano;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasXArtic;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasArtXMes;
		private DevExpress.XtraBars.BarButtonItem barButtonItem81;
		private DevExpress.XtraBars.BarSubItem barSubItem12;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasComisionesConf;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasComisionesCali;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasComisionesDetGe;
		private DevExpress.XtraBars.BarButtonItem mnuRPTVentasComisionesConsulta;
		private DevExpress.XtraBars.BarButtonItem mnuCompraAODC;
		private DevExpress.XtraBars.BarButtonItem barButtonItem83;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobCM;
		private DevExpress.XtraBars.BarButtonItem mnuReporVentasPDGA;
		private DevExpress.XtraBars.BarButtonItem mnuCRResumenCredito;
		private DevExpress.XtraBars.BarButtonItem barButtonItem82;
		private DevExpress.XtraBars.BarButtonItem mnuCom_ReporteLM;
		private DevExpress.XtraBars.BarButtonItem mnuBodegaNI;
		private DevExpress.XtraBars.BarButtonItem mnuComprasLAC;
		private DevExpress.XtraBars.BarSubItem mnuCPComisiones;
		private DevExpress.XtraBars.BarButtonItem mnuCPComisionesAsigCobr;
		private DevExpress.XtraBars.BarButtonItem mnuCPComisionesTrmOpe;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobranzasCboOpe;
		private DevExpress.XtraBars.BarButtonItem mnuCPComisionesCom;
		private DevExpress.XtraBars.BarButtonItem mnuCreditoGesCall;
		private DevExpress.XtraBars.BarButtonItem mnuCreditoProReal;
		private DevExpress.XtraBars.BarButtonItem mnuRPPagoAgil;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobranzasCERTIFICADO;
		private DevExpress.XtraBars.BarButtonItem mnuRPtComprasPuC;
		private DevExpress.XtraBars.BarButtonItem barButtonItem84;
		private DevExpress.XtraBars.BarButtonItem mnuCPANDC;
		private DevExpress.XtraBars.BarButtonItem mnuComprasPromoFact;
		private DevExpress.XtraBars.BarButtonItem barButtonItem85;
		private DevExpress.XtraBars.BarButtonItem mnuSolNC;
		private DevExpress.XtraBars.BarButtonItem mnuRolPorCajero;
		private DevExpress.XtraBars.BarButtonItem barButtonItem86;
		private DevExpress.XtraBars.BarButtonItem barButtonItem87;
		private DevExpress.XtraBars.BarButtonItem barButtonItem88;
		private DevExpress.XtraBars.BarButtonItem mnuRPtCREPOINTFDG;
		private DevExpress.XtraBars.BarButtonItem mnuRPtCREPOINRMD;
		private DevExpress.XtraBars.BarButtonItem mnuCPCobranzasRPP;
		private DevExpress.XtraBars.BarButtonItem mnuReporVentasDFI;
		private DevExpress.XtraBars.BarButtonItem mnuRPtContabilidadSCosecha;
		private DevExpress.XtraBars.BarButtonItem mnuContLotAsiComprob;
		#endregion mnulat

		private void mnuFacRepSri(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ReporteSRI miRep = new ReporteSRI();
			miRep.MdiParent = this;
			miRep.Show();
		}
	
		private DevExpress.XtraBars.BarButtonItem MnuFactLot_Autoriza;
		private DevExpress.XtraBars.BarButtonItem mneArchImpClaseEcuador;
		private DevExpress.XtraBars.BarButtonItem mnuProdLiquida;
		private DevExpress.XtraBars.BarButtonItem mneArchImpFidelius;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpFresh;
		private DevExpress.XtraBars.BarButtonItem mnuArchImpCash;
		private DevExpress.XtraBars.BarButtonItem mnuRolUtilidad;
		private DevExpress.XtraBars.BarButtonItem MnuArchPropNumE;
		private DevExpress.XtraBars.BarButtonItem mnuFactRepSri;
		private System.ComponentModel.IContainer components;

		public MenuLatinium()
		{
			InitializeComponent();
		}

		/// <summary>
		/// Limpiar los recursos que se estén utilizando.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MenuLatinium));
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.brAbrir = new DevExpress.XtraBars.BarButtonItem();
			this.brNuevo = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem1 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem2 = new DevExpress.XtraBars.BarButtonItem();
			this.lblTiempos = new DevExpress.XtraBars.BarManager(this.components);
			this.bar1 = new DevExpress.XtraBars.Bar();
			this.MnuArch = new DevExpress.XtraBars.BarSubItem();
			this.MnuArchNuevo = new DevExpress.XtraBars.BarButtonItem();
			this.MnuArchAbrir = new DevExpress.XtraBars.BarButtonItem();
			this.MnuArchGuardar = new DevExpress.XtraBars.BarButtonItem();
			this.MnuArchImpoExpo = new DevExpress.XtraBars.BarSubItem();
			this.MnuArchImpoLat = new DevExpress.XtraBars.BarButtonItem();
			this.MnuArchImpoOtros = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpOmniLife = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpTelcosan = new DevExpress.XtraBars.BarButtonItem();
			this.mneArchImpInterActive = new DevExpress.XtraBars.BarButtonItem();
			this.mneArchImpYellow = new DevExpress.XtraBars.BarButtonItem();
			this.mneArchImpClaseEcuador = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpFresh = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpCash = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpQBI = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpCashE = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpCerv = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpResCDI = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpCostoI = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpSeprofin = new DevExpress.XtraBars.BarButtonItem();
			this.barSubItem1 = new DevExpress.XtraBars.BarSubItem();
			this.mnuArchImpGuiaGruas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpGruasCompra = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpGuiaVenta = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpJuanMarcet = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpFactRet = new DevExpress.XtraBars.BarButtonItem();
			this.MnuArchProp = new DevExpress.XtraBars.BarSubItem();
			this.mnuArchPropGen = new DevExpress.XtraBars.BarButtonItem();
			this.MnuArchPropCont = new DevExpress.XtraBars.BarButtonItem();
			this.MnuArchPropFact = new DevExpress.XtraBars.BarButtonItem();
			this.MnuArchPropNumE = new DevExpress.XtraBars.BarButtonItem();
			this.MnuArchPropRol = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchPropProd = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchPropEmpresa = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchPropEmpresaR = new DevExpress.XtraBars.BarStaticItem();
			this.mnuArchPropConfig = new DevExpress.XtraBars.BarButtonItem();
			this.MnuArchConfImp = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem14 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem12 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem17 = new DevExpress.XtraBars.BarButtonItem();
			this.MnuArchSalir = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem53 = new DevExpress.XtraBars.BarButtonItem();
			this.MnuEdic = new DevExpress.XtraBars.BarSubItem();
			this.MnuEdicDesh = new DevExpress.XtraBars.BarButtonItem();
			this.MnuEdicReha = new DevExpress.XtraBars.BarButtonItem();
			this.MnuEdicCortar = new DevExpress.XtraBars.BarButtonItem();
			this.MnuEdicCopi = new DevExpress.XtraBars.BarButtonItem();
			this.MnuEdicPega = new DevExpress.XtraBars.BarButtonItem();
			this.MnuBodega = new DevExpress.XtraBars.BarSubItem();
			this.mnuBodegaTRM = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactSerie = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFacInv = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaKardex = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactLoteSpif = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFacSeriales = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaIngSeriales = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaIPR = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactTransf = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactLoteTransMult = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasTEL = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactEstTransf = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactGuiaRemision = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaGRC = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactEnvTransferencia = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactRecTransferencia = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaNI = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaCDI = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaRM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaCI = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaDCI = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaCierreInventario = new DevExpress.XtraBars.BarButtonItem();
			this.mnuReportesBodega = new DevExpress.XtraBars.BarSubItem();
			this.mnuBodegaIF = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaInvVal = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRepCmpTransf = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaRepKS = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaReportesECDI = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaRepMD = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaReportesKardexCosto = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactBodegas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaSectorizacion = new DevExpress.XtraBars.BarButtonItem();
			this.MnuHerr_ = new DevExpress.XtraBars.BarSubItem();
			this.MnuHerr_Camb_Contr = new DevExpress.XtraBars.BarButtonItem();
			this.MnuHerr_Actu_Vers = new DevExpress.XtraBars.BarButtonItem();
			this.MnuHerr_Sal_Ini = new DevExpress.XtraBars.BarButtonItem();
			this.MnuHerr_Camb_Ctes = new DevExpress.XtraBars.BarButtonItem();
			this.MnuHerr_Copi_Arch = new DevExpress.XtraBars.BarButtonItem();
			this.MnuHerr_Auditoría = new DevExpress.XtraBars.BarButtonItem();
			this.MnuHerr_Act_Saldo = new DevExpress.XtraBars.BarButtonItem();
			this.MnuEdicSeguridad = new DevExpress.XtraBars.BarSubItem();
			this.MnuEdicSegMenu = new DevExpress.XtraBars.BarButtonItem();
			this.MnuEdicSegUsuario = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrSegGenGrupos = new DevExpress.XtraBars.BarButtonItem();
			this.MnuEdicSegConfigUsuarios = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrSegFiltro = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrSegFiltroSuc = new DevExpress.XtraBars.BarButtonItem();
			this.mbuHerrSegTransf = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrSegApReq = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrSegReqLis = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrSegCD = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem68 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrGenSQL = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrTester = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrReporte = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactRepComp = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrCredito = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrBusqueda = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrContratos = new DevExpress.XtraBars.BarSubItem();
			this.mnuHerrContratosIngreso = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrContratosGeneraFact = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrContratosAmortiza = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrContratosAdemdums = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrContratosAlarma = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrPresupC = new DevExpress.XtraBars.BarSubItem();
			this.mnuHerrPresupCIngreso = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrPresupCAvanceObra = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrPresupCSolicitudPago = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrPresupCAnalisis = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrPresupCReportes = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrPresupCImportar = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerramientasNumeracion = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerramientasCRIVA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrALF = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerramientasPTP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerramientasAP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerramientasPTA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerramientasAPL = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerramientasASDP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasEOC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasGCR = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrVerifDepositos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrSegClientes = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrDepositos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuSolNC = new DevExpress.XtraBars.BarButtonItem();
			this.MnuCont = new DevExpress.XtraBars.BarSubItem();
			this.MnuContPlanCtas = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContAsigCtas = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContAsientos = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFacArtServicios = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactProyectos = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContFlujo = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContTCC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContConsumosTC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContLCTC = new DevExpress.XtraBars.BarButtonItem();
			this.barSubItem4 = new DevExpress.XtraBars.BarSubItem();
			this.MnuContBcos = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContCtas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContMovBanc = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContEstadoCheques = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContTEB = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContPrestamos = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContEstadoCta = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContBancosCBA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContBancosCP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContBancosDPL = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContBancosCDPL = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContSRI = new DevExpress.XtraBars.BarSubItem();
			this.mnuContTransAnexoTr = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContTransFormulario = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContTransConciliacion = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactAutPagos = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFacPagosAut = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactActFijos = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContBloqAsientos = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContCierPeriodo = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContReportes = new DevExpress.XtraBars.BarSubItem();
			this.MnuContRepCont = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContabilidadAsDesc = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContTransR = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContRepCajasChicas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContRepGastos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContReportesGA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContRepFacElec = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContRepEstadoCB = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContRepPresupAlm = new DevExpress.XtraBars.BarButtonItem();
			this.MnuCont_Ind_Financieros = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContPresupuesto = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContLote = new DevExpress.XtraBars.BarSubItem();
			this.mnuContCompara = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContAsientoPeriodo = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContLoteDepura = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContLotAsiComprob = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContLotCuentaObligatoria = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContTransacciones = new DevExpress.XtraBars.BarSubItem();
			this.MnuContTransE = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContTransI = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContRepCC = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContAsigArt = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContabilidadProvisiones = new DevExpress.XtraBars.BarSubItem();
			this.mnuContabilidadFlujoPMP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContabilidadFPPP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContFPPP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContRB = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContCaja = new DevExpress.XtraBars.BarSubItem();
			this.mnuContEstadosCajas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContIngresoDiarioCaja = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContIngresosDeCajas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContEstadoGeneralDeCaja = new DevExpress.XtraBars.BarButtonItem();
			this.MnuRol = new DevExpress.XtraBars.BarSubItem();
			this.MnuRolPersonal = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolPA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolPersonalPrueba = new DevExpress.XtraBars.BarButtonItem();
			this.MnuRolCargos = new DevExpress.XtraBars.BarButtonItem();
			this.MnuRolDpto = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesComisiones = new DevExpress.XtraBars.BarSubItem();
			this.mnuRolComPIA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesCom_PA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCom_PAGE = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesComisionesPS = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesCom_PL = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesCom_PML = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCom_PMLGE = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesCom_AV = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesComAJV = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolComisionesPC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolCom_PorcCarpa = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCom_PCGE = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesCom_EPV = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCom_Reportes = new DevExpress.XtraBars.BarSubItem();
			this.mnuCom_ReportesRM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCom_ReporteLM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolPorCajero = new DevExpress.XtraBars.BarButtonItem();
			this.MnuRolVerif = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolFiniquitos = new DevExpress.XtraBars.BarButtonItem();
			this.MnuRolGenRol = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesPagoComisiones = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesDecimoTercerSueldo = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesReportes = new DevExpress.XtraBars.BarSubItem();
			this.mnuRolesReportesPVF = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesReportesVXM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesReportesUniformes = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesReportesNominaGeneral = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolReportesAE = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesReportesHHT = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolReportesIND = new DevExpress.XtraBars.BarButtonItem();
			this.barRolRubro = new DevExpress.XtraBars.BarButtonItem();
			this.MnuRolRepRol = new DevExpress.XtraBars.BarButtonItem();
			this.MnuRolBloqRol = new DevExpress.XtraBars.BarButtonItem();
			this.MnuRolIncSuel = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesLiquidacion = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolUtilidad = new DevExpress.XtraBars.BarButtonItem();
			this.MnuRolAsigCtas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolAnexo = new DevExpress.XtraBars.BarButtonItem();
			this.mnuTarea_Agr = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesDistribucion = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesAdicional = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesAdicionalFormulas = new DevExpress.XtraBars.BarButtonItem();
			this.MnuCompras = new DevExpress.XtraBars.BarSubItem();
			this.MnuFactProvee = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactOrdComp = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCompraAODC = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactCompras = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactDevCompras = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactNotaCC = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactLotimp = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasLAC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactVentasConsignacion = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContAntProv = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContASF = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContCXP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContResCXP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasGDA = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactArtic = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCompraCSA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesVA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasCatalogo = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactCombos = new DevExpress.XtraBars.BarButtonItem();
			this.MnnuFactLotesPromociones = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasPromociones = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactActPrecio = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactLotesPromVen = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasRepSpifVend = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasAgenda = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCompraGestion = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasLiquidaGastos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasComboZOOM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuMatrizDsctos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasReportes = new DevExpress.XtraBars.BarSubItem();
			this.mnuCompraReportesCPA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasAPC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasRepFP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasRepCodigoArticulos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasProformas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuComprasPromoFact = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentas = new DevExpress.XtraBars.BarSubItem();
			this.MnuFactClientes = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactVentas = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem66 = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactPreFactura = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasPFR = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactPedClientes = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactReservClientes = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasCotizacion = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactConsignacion = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentaLiqCons = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasECW = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactDevVentas = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactNotaCV = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactAnticipos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasDA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasCNOS = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasRPA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentaAutDistribuidor = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasCB = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContCuadreCaja = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContDepositoVouchers = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContMantTarj = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasResumenDiario = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasResumenComisiones = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasMarcador = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasPlanifVentas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasPlanifJefeVentas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasPC = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContCXC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasImpresionDePrecios = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasIPP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasIPR = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasARP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasRP = new DevExpress.XtraBars.BarButtonItem();
			this.barSubItem3 = new DevExpress.XtraBars.BarSubItem();
			this.mnuVentasReporteCompara = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesPorMarca = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasRVML = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasRepMU = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentaRMU = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentaRMUA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesAIM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesRF = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesVPD = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesFFP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasRepGE = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasRepRVF = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesComparativos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesVentasPorMes = new DevExpress.XtraBars.BarButtonItem();
			this.mnuReporteVentasMensualesPorLocal = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesTC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesDTC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesEncuestas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesCXC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContResCXC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesVPA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasComisiones = new DevExpress.XtraBars.BarSubItem();
			this.MnuFacComConfiguracion = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactComCalificacion = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactDetalle = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactComConsulta = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesVentasPorArticulos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasReportesVentasarticuloPorMes = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentaPromocionesActivas = new DevExpress.XtraBars.BarButtonItem();
			this.MnuCredito = new DevExpress.XtraBars.BarSubItem();
			this.MnuCreCobranza = new DevExpress.XtraBars.BarButtonItem();
			this.MnuCreditoCotizador = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPEDC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPHRC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPRRC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasGCCC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPRevision = new DevExpress.XtraBars.BarButtonItem();
			this.MnuCreditoSolRapida = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCrediPointSolicitud = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCrediPointMEF = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPMFCP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCrediPointMantenedor = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCrediPointVerTerreno = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCrediPointGestiones = new DevExpress.XtraBars.BarSubItem();
			this.mnuCrediPointGestVerCed = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCrediPointGestRevCed = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem51 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCrediPointCatalogos = new DevExpress.XtraBars.BarSubItem();
			this.MnuCreditoCuota = new DevExpress.XtraBars.BarButtonItem();
			this.MnuCreGastoCobran = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPReportes = new DevExpress.XtraBars.BarSubItem();
			this.mnuCPReportesCliPoint = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPLCS = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPReportesEC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPRepEquifax = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPRepCP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPRepRDD = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPRepECP = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem37 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPReportesRCV = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPReportesGDCM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPReportesAbonos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPReportesNDC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPReportesBR = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem76 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPReportesREC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCrediPOINTReportesRMR = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPReportesRRM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPReportesCVPM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCRResumenCredito = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobranzas = new DevExpress.XtraBars.BarSubItem();
			this.mnuCPCobranzasGDC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobranzasCboOpe = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPInformeCobrador = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPIO = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCobranzasRGC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobranzasRCP = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem50 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobranzasCCPPC = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem72 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobranzasLPCA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobranzasEDC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobranzasMRPR = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobranzasRPE = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobCM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobranzasRPP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPComisiones = new DevExpress.XtraBars.BarSubItem();
			this.mnuCPComisionesAsigCobr = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPComisionesTrmOpe = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPComisionesCom = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCreditoGesCall = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCreditoProReal = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem84 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCartera = new DevExpress.XtraBars.BarSubItem();
			this.barSubItem10 = new DevExpress.XtraBars.BarSubItem();
			this.mnuCrediPointDFG = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCreLotesUnibanco = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContRepAnt = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPRRCS = new DevExpress.XtraBars.BarButtonItem();
			this.mnuContCCxCCP = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem57 = new DevExpress.XtraBars.BarButtonItem();
			this.barSubItem11 = new DevExpress.XtraBars.BarSubItem();
			this.mnuCpLbp = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPLotesPPagos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPDFGP = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem82 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobranzasCarteraAIP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCobranzasCERTIFICADO = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPRD = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPVEP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCrediPOINTRMR = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCRD = new DevExpress.XtraBars.BarSubItem();
			this.mnuCPCDRDC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCDRVC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCDSR = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCDRSLT = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPCDLCPG = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCPANDC = new DevExpress.XtraBars.BarButtonItem();
			this.MnuProduc = new DevExpress.XtraBars.BarSubItem();
			this.MnuProducRecetas = new DevExpress.XtraBars.BarButtonItem();
			this.MnuProducOrdProd = new DevExpress.XtraBars.BarButtonItem();
			this.MnuProdEstado = new DevExpress.XtraBars.BarButtonItem();
			this.mnuProduccionPendientes = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem18 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuProduccionTPT = new DevExpress.XtraBars.BarButtonItem();
			this.mnuGPPD = new DevExpress.XtraBars.BarButtonItem();
			this.MnuProducProdTerm = new DevExpress.XtraBars.BarButtonItem();
			this.mnuProdLiquida = new DevExpress.XtraBars.BarButtonItem();
			this.MnuProducConsReal = new DevExpress.XtraBars.BarButtonItem();
			this.MnuProducActualiza = new DevExpress.XtraBars.BarButtonItem();
			this.MnuProducReportes = new DevExpress.XtraBars.BarButtonItem();
			this.MnuProdClientes = new DevExpress.XtraBars.BarButtonItem();
			this.MnuProdSemana = new DevExpress.XtraBars.BarButtonItem();
			this.btProdCosto = new DevExpress.XtraBars.BarButtonItem();
			this.mnuProdGeneraC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuProdAgro = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDT = new DevExpress.XtraBars.BarSubItem();
			this.MnuDTSolicitud = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIREP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTRA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIREPC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIngresoMultiple = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTRetiro = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRMP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTProductosDeFeria = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRAT = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRSNC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTAprobacionRemates = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRRPT = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRRMA = new DevExpress.XtraBars.BarSubItem();
			this.mnuDTIRRPG = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRRSR = new DevExpress.XtraBars.BarButtonItem();
			this.MnuDTSTA = new DevExpress.XtraBars.BarButtonItem();
			this.frmDTTransfLocales = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRAPD = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTASSF = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRREP = new DevExpress.XtraBars.BarSubItem();
			this.mnuDTReporteEstadoIngresos = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRRRP = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRRTE = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRFDT = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRSIN = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTSubGrupo = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRCDE = new DevExpress.XtraBars.BarButtonItem();
			this.mnuLotte = new DevExpress.XtraBars.BarSubItem();
			this.mnuCuadroHabitaciones = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem26 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuPSRDC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCobrPrenda = new DevExpress.XtraBars.BarButtonItem();
			this.mnuReservas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuPromociones = new DevExpress.XtraBars.BarButtonItem();
			this.mnuKardexLot = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem27 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem28 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem29 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuLoteNotasDeEntrega = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHabitaciones = new DevExpress.XtraBars.BarButtonItem();
			this.mnuTmaHabitacion = new DevExpress.XtraBars.BarButtonItem();
			this.mnuTmpEstadia = new DevExpress.XtraBars.BarButtonItem();
			this.mnuTmpHabitaciones = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem36 = new DevExpress.XtraBars.BarButtonItem();
			this.barSubItem9 = new DevExpress.XtraBars.BarSubItem();
			this.barButtonItem24 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem25 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem31 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem32 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem22 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem35 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem40 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem42 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem67 = new DevExpress.XtraBars.BarButtonItem();
			this.barSubItem14 = new DevExpress.XtraBars.BarSubItem();
			this.mnuLoteRepProcSIM = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem69 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem75 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuLotteReportesIT = new DevExpress.XtraBars.BarButtonItem();
			this.mnuLotteReporteIPH = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem30 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuLottePlanificacion = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem73 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuLotteCompetencia = new DevExpress.XtraBars.BarButtonItem();
			this.mnuTarjetasDeIngreso = new DevExpress.XtraBars.BarButtonItem();
			this.MnuVentana = new DevExpress.XtraBars.BarSubItem();
			this.MnuVentana_Casc = new DevExpress.XtraBars.BarButtonItem();
			this.MnuVentana_Org = new DevExpress.XtraBars.BarButtonItem();
			this.MnuVentana_Cicl = new DevExpress.XtraBars.BarButtonItem();
			this.MnuVentTrans = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentOpac = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFact = new DevExpress.XtraBars.BarSubItem();
			this.MnuFactCotizacion = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactResPed = new DevExpress.XtraBars.BarSubItem();
			this.MnuFactReservProveedores = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactResPedProv = new DevExpress.XtraBars.BarButtonItem();
			this.mnuResPedPend = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactPreciosProveedores = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactResPedRequisicion = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactResPedComision = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactResCruzarFactReserva = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactResOTrabajo = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactResConclusion = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactResOTrabajoServicio = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactResJarComision = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactLotes = new DevExpress.XtraBars.BarSubItem();
			this.MnuFactLot_Ctas_Cob = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactLot_Ctas_Pag = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactLot_Autoriza = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactLot_Facturacion = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactLotesPago = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactLot_Transf_Art = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFacLotTransfTrans = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactLoteAnalisis = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactLotAgencia = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactLotTomaF = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFacLotCxC30 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactLoteTarjeta = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFacLotSerie = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactRepSri = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFacLotCierre = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFacLotCopiaArt = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFacLotFactReembol = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactGrupArticulos = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFacSerie = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactCtasCobrar = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactCtasPagar = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactRepResumenes = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactRepInventario = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactRepArticulo = new DevExpress.XtraBars.BarButtonItem();
			this.mnuMetricas = new DevExpress.XtraBars.BarSubItem();
			this.mnuMetricasRRHH = new DevExpress.XtraBars.BarSubItem();
			this.barButtonItem60 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem62 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuPCGEx = new DevExpress.XtraBars.BarButtonItem();
			this.mnuSTA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDirectorio = new DevExpress.XtraBars.BarButtonItem();
			this.barSubItem5 = new DevExpress.XtraBars.BarSubItem();
			this.mnuRPtBodega = new DevExpress.XtraBars.BarSubItem();
			this.mnuRPtBodegaInvFEcha = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtBodegaRptTrans = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtBodegaInvVal = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtBodegaKXSer = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtBodegaEsConInv = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtBodegaMovDesc = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtBodegaKCC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtContabilidad = new DevExpress.XtraBars.BarSubItem();
			this.mnuRPtContabilidadCont = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtContabilidadRet = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtContabilidadAsDEsc = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtContabilidadGast = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtContabilidadFactEl = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtContabilidadCaCh = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtContabilidadGaAn = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtContabilidadEstaConB = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtContabilidadPrXAlm = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPPagoAgil = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtContabilidadSCosecha = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtRoles = new DevExpress.XtraBars.BarSubItem();
			this.mnuRPtRolesPFA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtRolesVeXM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtRolesUni = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtRolesNG = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtRolesAEM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtRolesHHT = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtRolesIXD = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtCompras = new DevExpress.XtraBars.BarSubItem();
			this.mnuRPtComprasCPA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtComprasT = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtComprasRAC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtComprasCA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtComprasPuC = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtVentas = new DevExpress.XtraBars.BarSubItem();
			this.mnuRPTVentasPorcMarca = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasPorcMarcaUtilid = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasMargUti = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasReMarUti = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasMaUXArt = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasAnInMEn = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasRevFact = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasXDia = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasNFactPForPago = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasGaraExt = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasXFecha = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasComDeVentas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasANuXLo = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasMenXLo = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasTarCredi = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasDetTc = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasEncuestas = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasReCXCo = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasXano = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasXArtic = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasArtXMes = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasPorcMargUtilidad = new DevExpress.XtraBars.BarButtonItem();
			this.barSubItem12 = new DevExpress.XtraBars.BarSubItem();
			this.mnuRPTVentasComisionesConf = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasComisionesCali = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasComisionesDetGe = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPTVentasComisionesConsulta = new DevExpress.XtraBars.BarButtonItem();
			this.mnuReporVentasPDGA = new DevExpress.XtraBars.BarButtonItem();
			this.mnuReporVentasDFI = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtCREPOINT = new DevExpress.XtraBars.BarSubItem();
			this.mnuRPtCREPOINTFDG = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtCREPOINRMD = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRPtDT = new DevExpress.XtraBars.BarSubItem();
			this.mnuRPtPeonyServicios = new DevExpress.XtraBars.BarSubItem();
			this.mnuIngreso = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem86 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem87 = new DevExpress.XtraBars.BarButtonItem();
			this.bar2 = new DevExpress.XtraBars.Bar();
			this.barNuevo = new DevExpress.XtraBars.BarButtonItem();
			this.barAbrir = new DevExpress.XtraBars.BarButtonItem();
			this.barGrabar = new DevExpress.XtraBars.BarButtonItem();
			this.BarPlanCuentas = new DevExpress.XtraBars.BarButtonItem();
			this.barAsiento = new DevExpress.XtraBars.BarButtonItem();
			this.BarRepContabilidad = new DevExpress.XtraBars.BarButtonItem();
			this.BarCompras = new DevExpress.XtraBars.BarButtonItem();
			this.BarVentas = new DevExpress.XtraBars.BarButtonItem();
			this.BarArticulo = new DevExpress.XtraBars.BarButtonItem();
			this.BarCliente = new DevExpress.XtraBars.BarButtonItem();
			this.BarRepResumen = new DevExpress.XtraBars.BarButtonItem();
			this.BarRepInventario = new DevExpress.XtraBars.BarButtonItem();
			this.BarRepArticulo = new DevExpress.XtraBars.BarButtonItem();
			this.barSeguridad = new DevExpress.XtraBars.BarButtonItem();
			this.barPagar = new DevExpress.XtraBars.BarButtonItem();
			this.barCobro = new DevExpress.XtraBars.BarButtonItem();
			this.lblMensaje = new DevExpress.XtraBars.BarEditItem();
			this.repositoryItemTextEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
			this.barSkin = new DevExpress.XtraBars.BarEditItem();
			this.repositoryItemComboBox1 = new DevExpress.XtraEditors.Repository.RepositoryItemComboBox();
			this.barDockControlTop = new DevExpress.XtraBars.BarDockControl();
			this.barDockControlBottom = new DevExpress.XtraBars.BarDockControl();
			this.barDockControlLeft = new DevExpress.XtraBars.BarDockControl();
			this.barDockControlRight = new DevExpress.XtraBars.BarDockControl();
			this.MnuHerr_Cre_Sal_Ini = new DevExpress.XtraBars.BarButtonItem();
			this.MnuHerr_Cop_Sal_Ini = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContRepGov = new DevExpress.XtraBars.BarButtonItem();
			this.MnuContFormulario = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactPedEmpleados = new DevExpress.XtraBars.BarButtonItem();
			this.mnuFactLotOrdenMov = new DevExpress.XtraBars.BarButtonItem();
			this.MnuRolDatGen = new DevExpress.XtraBars.BarButtonItem();
			this.MnuProducPrecios = new DevExpress.XtraBars.BarButtonItem();
			this.MnuAyuda = new DevExpress.XtraBars.BarSubItem();
			this.MnuAyuda_Ayuda = new DevExpress.XtraBars.BarButtonItem();
			this.MnuAyudaRegist = new DevExpress.XtraBars.BarButtonItem();
			this.MnuAyudaAcercaCIAD = new DevExpress.XtraBars.BarButtonItem();
			this.mnuAyuLiberar = new DevExpress.XtraBars.BarButtonItem();
			this.btControlRemoto = new DevExpress.XtraBars.BarButtonItem();
			this.mnuAyuJuego = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem63 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrSegReqEst = new DevExpress.XtraBars.BarButtonItem();
			this.mnuProdOrdenT = new DevExpress.XtraBars.BarButtonItem();
			this.mneArchImpFidelius = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrSegTransf = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrContratosAdemdum = new DevExpress.XtraBars.BarButtonItem();
			this.mnuArchImpMarceloVentas = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactLotimp2 = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFactContometros = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem3 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem4 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuHerrAgenda = new DevExpress.XtraBars.BarButtonItem();
			this.MnuDepTec = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem5 = new DevExpress.XtraBars.BarButtonItem();
			this.MnuHerrAlarma = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFacSpif = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem6 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem11 = new DevExpress.XtraBars.BarButtonItem();
			this.MnuFacOCImp = new DevExpress.XtraBars.BarButtonItem();
			this.barSubItem6 = new DevExpress.XtraBars.BarSubItem();
			this.barSubItem7 = new DevExpress.XtraBars.BarSubItem();
			this.FacGuiaRemision = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem16 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasComDetallePersonal = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem7 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem8 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem10 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaReportes = new DevExpress.XtraBars.BarButtonItem();
			this.mnuTomaInventario = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem13 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem9 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuCmpRepVM = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem19 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem20 = new DevExpress.XtraBars.BarButtonItem();
			this.barToolbarsListItem1 = new DevExpress.XtraBars.BarToolbarsListItem();
			this.barButtonItem21 = new DevExpress.XtraBars.BarButtonItem();
			this.barSubItem8 = new DevExpress.XtraBars.BarSubItem();
			this.barButtonItem23 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVentasVentasCP = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem33 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem34 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem38 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem39 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem41 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem44 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem45 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem46 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem47 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem48 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem49 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem55 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem56 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuDTIRPPL = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem61 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem70 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem71 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuVRE = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem74 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem58 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem59 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem52 = new DevExpress.XtraBars.BarButtonItem();
			this.mnuBodegaTIM = new DevExpress.XtraBars.BarButtonItem();
			this.mnuRolesComisionesPCS = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem43 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem54 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem64 = new DevExpress.XtraBars.BarButtonItem();
			this.barLinkContainerItem1 = new DevExpress.XtraBars.BarLinkContainerItem();
			this.barButtonItem65 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem77 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem78 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem79 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem80 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem81 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem83 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem85 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem88 = new DevExpress.XtraBars.BarButtonItem();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.helpProvider1 = new System.Windows.Forms.HelpProvider();
			this.timer = new System.Windows.Forms.Timer(this.components);
			this.skinPrincipal = new DevExpress.LookAndFeel.DefaultLookAndFeel(this.components);
			this.barSubItem2 = new DevExpress.XtraBars.BarSubItem();
			this.barButtonItem15 = new DevExpress.XtraBars.BarButtonItem();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			((System.ComponentModel.ISupportInitialize)(this.lblTiempos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemComboBox1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			this.SuspendLayout();
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// brAbrir
			// 
			this.brAbrir.Caption = "Abrir";
			this.brAbrir.Id = 5;
			this.brAbrir.ImageIndex = 5;
			this.brAbrir.Name = "brAbrir";
			// 
			// brNuevo
			// 
			this.brNuevo.Caption = "Nuevo";
			this.brNuevo.Id = 4;
			this.brNuevo.ImageIndex = 0;
			this.brNuevo.Name = "brNuevo";
			// 
			// barButtonItem1
			// 
			this.barButtonItem1.Caption = "Nuevo";
			this.barButtonItem1.Id = 0;
			this.barButtonItem1.ImageIndex = 0;
			this.barButtonItem1.Name = "barButtonItem1";
			// 
			// barButtonItem2
			// 
			this.barButtonItem2.Caption = "Abrir";
			this.barButtonItem2.Id = 2;
			this.barButtonItem2.ImageIndex = 5;
			this.barButtonItem2.Name = "barButtonItem2";
			// 
			// lblTiempos
			// 
			this.lblTiempos.Bars.AddRange(new DevExpress.XtraBars.Bar[] {
																																		this.bar1,
																																		this.bar2});
			this.lblTiempos.Categories.AddRange(new DevExpress.XtraBars.BarManagerCategory[] {
																																												 new DevExpress.XtraBars.BarManagerCategory("Archivo", new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759")),
																																												 new DevExpress.XtraBars.BarManagerCategory("Herramientas", new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c")),
																																												 new DevExpress.XtraBars.BarManagerCategory("Contabilidad", new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3")),
																																												 new DevExpress.XtraBars.BarManagerCategory("Facturacion", new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9")),
																																												 new DevExpress.XtraBars.BarManagerCategory("Roles", new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b")),
																																												 new DevExpress.XtraBars.BarManagerCategory("Produccion", new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9")),
																																												 new DevExpress.XtraBars.BarManagerCategory("Edicion", new System.Guid("d9189416-8661-46d9-99d3-d809add0384f")),
																																												 new DevExpress.XtraBars.BarManagerCategory("Ventana", new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88"))});
			this.lblTiempos.DockControls.Add(this.barDockControlTop);
			this.lblTiempos.DockControls.Add(this.barDockControlBottom);
			this.lblTiempos.DockControls.Add(this.barDockControlLeft);
			this.lblTiempos.DockControls.Add(this.barDockControlRight);
			this.lblTiempos.Form = this;
			this.lblTiempos.Images = this.imageList1;
			this.lblTiempos.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
																																				 this.MnuArch,
																																				 this.MnuEdic,
																																				 this.MnuArchNuevo,
																																				 this.barNuevo,
																																				 this.MnuArchAbrir,
																																				 this.barAbrir,
																																				 this.MnuHerr_,
																																				 this.MnuArchGuardar,
																																				 this.barGrabar,
																																				 this.MnuArchImpoExpo,
																																				 this.MnuArchImpoLat,
																																				 this.MnuArchImpoOtros,
																																				 this.MnuArchProp,
																																				 this.MnuCont,
																																				 this.MnuContPlanCtas,
																																				 this.BarPlanCuentas,
																																				 this.MnuContAsientos,
																																				 this.barAsiento,
																																				 this.mnuArchPropEmpresa,
																																				 this.MnuArchPropCont,
																																				 this.MnuArchPropFact,
																																				 this.MnuArchPropRol,
																																				 this.mnuArchPropProd,
																																				 this.mnuArchPropGen,
																																				 this.MnuArchPropNumE,
																																				 this.mnuArchPropConfig,
																																				 this.MnuEdicDesh,
																																				 this.MnuEdicReha,
																																				 this.MnuEdicCortar,
																																				 this.MnuEdicCopi,
																																				 this.MnuEdicPega,
																																				 this.MnuHerr_Camb_Contr,
																																				 this.MnuHerr_Actu_Vers,
																																				 this.MnuHerr_Sal_Ini,
																																				 this.MnuHerr_Camb_Ctes,
																																				 this.MnuHerr_Copi_Arch,
																																				 this.MnuHerr_Auditoría,
																																				 this.MnuHerr_Act_Saldo,
																																				 this.MnuEdicSeguridad,
																																				 this.MnuEdicSegMenu,
																																				 this.MnuEdicSegUsuario,
																																				 this.MnuHerr_Cre_Sal_Ini,
																																				 this.MnuHerr_Cop_Sal_Ini,
																																				 this.MnuContRepCont,
																																				 this.BarRepContabilidad,
																																				 this.MnuContRepGov,
																																				 this.MnuContTransacciones,
																																				 this.MnuContTransE,
																																				 this.MnuContTransI,
																																				 this.MnuContTransR,
																																				 this.mnuContTransAnexoTr,
																																				 this.mnuContTransFormulario,
																																				 this.MnuContEstadoCta,
																																				 this.MnuContPresupuesto,
																																				 this.MnuContCtas,
																																				 this.MnuContBcos,
																																				 this.MnuContBloqAsientos,
																																				 this.MnuContCierPeriodo,
																																				 this.MnuContAsigCtas,
																																				 this.MnuContFormulario,
																																				 this.MnuCont_Ind_Financieros,
																																				 this.MnuFact,
																																				 this.MnuFactCompras,
																																				 this.BarCompras,
																																				 this.MnuFactOrdComp,
																																				 this.MnuFactCotizacion,
																																				 this.MnuFactVentas,
																																				 this.BarVentas,
																																				 this.MnuFactDevVentas,
																																				 this.MnuFactDevCompras,
																																				 this.MnuFactEnvTransferencia,
																																				 this.MnuFactRecTransferencia,
																																				 this.mnuFactTransf,
																																				 this.MnuFactResPed,
																																				 this.MnuFactResPedProv,
																																				 this.MnuFactReservClientes,
																																				 this.MnuFactReservProveedores,
																																				 this.MnuFactPedEmpleados,
																																				 this.MnuFactPedClientes,
																																				 this.MnuFactPreciosProveedores,
																																				 this.MnuFactResPedRequisicion,
																																				 this.MnuFactResPedComision,
																																				 this.MnuFactNotaCV,
																																				 this.MnuFactNotaCC,
																																				 this.MnuFactLotes,
																																				 this.mnuResPedPend,
																																				 this.MnuFactLoteAnalisis,
																																				 this.MnuFactLot_Ctas_Cob,
																																				 this.MnuFactLot_Ctas_Pag,
																																				 this.mnuFactLotAgencia,
																																				 this.MnuFactLotimp,
																																				 this.mnuFactLotOrdenMov,
																																				 this.mnuFactLotTomaF,
																																				 this.mnuFactLoteTarjeta,
																																				 this.mnuFacLotCxC30,
																																				 this.MnuFactLot_Autoriza,
																																				 this.mnuFacLotSerie,
																																				 this.mnuFactRepSri,
																																				 this.mnuFacLotCierre,
																																				 this.MnuFactLot_Facturacion,
																																				 this.mnuFactLotesPago,
																																				 this.MnuFactLot_Transf_Art,
																																				 this.MnuFactArtic,
																																				 this.BarArticulo,
																																				 this.MnuFactClientes,
																																				 this.BarCliente,
																																				 this.MnuFactProvee,
																																				 this.MnuFactGrupArticulos,
																																				 this.MnuFactActFijos,
																																				 this.MnuFactProyectos,
																																				 this.MnuFactCtasCobrar,
																																				 this.barCobro,
																																				 this.MnuFactCtasPagar,
																																				 this.barPagar,
																																				 this.MnuFactRepResumenes,
																																				 this.BarRepResumen,
																																				 this.MnuFactRepInventario,
																																				 this.BarRepInventario,
																																				 this.MnuFactRepArticulo,
																																				 this.BarRepArticulo,
																																				 this.MnuRol,
																																				 this.MnuRolPersonal,
																																				 this.MnuRolCargos,
																																				 this.MnuRolDpto,
																																				 this.MnuRolGenRol,
																																				 this.MnuRolDatGen,
																																				 this.MnuRolRepRol,
																																				 this.MnuRolIncSuel,
																																				 this.MnuRolVerif,
																																				 this.MnuRolBloqRol,
																																				 this.MnuRolAsigCtas,
																																				 this.MnuProduc,
																																				 this.MnuProducRecetas,
																																				 this.MnuProducOrdProd,
																																				 this.MnuProducProdTerm,
																																				 this.MnuProducConsReal,
																																				 this.MnuProducActualiza,
																																				 this.MnuProducPrecios,
																																				 this.MnuProducReportes,
																																				 this.MnuVentana,
																																				 this.MnuAyuda,
																																				 this.MnuVentana_Casc,
																																				 this.MnuVentana_Org,
																																				 this.MnuVentana_Cicl,
																																				 this.MnuAyuda_Ayuda,
																																				 this.MnuAyudaRegist,
																																				 this.MnuAyudaAcercaCIAD,
																																				 this.MnuProdClientes,
																																				 this.MnuProdSemana,
																																				 this.barSeguridad,
																																				 this.mnuHerrSegFiltro,
																																				 this.mnuHerrSegFiltroSuc,
																																				 this.mbuHerrSegTransf,
																																				 this.mnuFactBodegas,
																																				 this.mnuHerrSegApReq,
																																				 this.mnuHerrSegReqLis,
																																				 this.mnuHerrSegReqEst,
																																				 this.mnuHerrGenSQL,
																																				 this.mnuHerrReporte,
																																				 this.barRolRubro,
																																				 this.mnuContLote,
																																				 this.mnuContCompara,
																																				 this.mnuContAsientoPeriodo,
																																				 this.mnuContLoteDepura,
																																				 this.MnuVentTrans,
																																				 this.mnuVentOpac,
																																				 this.mnuArchImpOmniLife,
																																				 this.mnuProdOrdenT,
																																				 this.mnuAyuLiberar,
																																				 this.mnuArchImpTelcosan,
																																				 this.mnuRolesLiquidacion,
																																				 this.mnuHerrTester,
																																				 this.mneArchImpInterActive,
																																				 this.mnuRolAnexo,
																																				 this.mneArchImpYellow,
																																				 this.mneArchImpClaseEcuador,
																																				 this.mnuProdLiquida,
																																				 this.mneArchImpFidelius,
																																				 this.mnuArchImpFresh,
																																				 this.mnuArchImpCash,
																																				 this.mnuRolUtilidad,
																																				 this.btProdCosto,
																																				 this.mnuArchImpGuiaGruas,
																																				 this.mnuArchImpGruasCompra,
																																				 this.mnuArchImpGuiaVenta,
																																				 this.mnuProdGeneraC,
																																				 this.mnuProdAgro,
																																				 this.mnuTarea_Agr,
																																				 this.mnuHerrSegTransf,
																																				 this.mnuArchImpQBI,
																																				 this.MnuArchConfImp,
																																				 this.MnuArchSalir,
																																				 this.mnuFactRepComp,
																																				 this.mnuHerrCredito,
																																				 this.mnuHerrBusqueda,
																																				 this.mnuContLotAsiComprob,
																																				 this.mnuArchImpCashE,
																																				 this.mnuArchImpCerv,
																																				 this.MnuContFlujo,
																																				 this.mnuArchImpResCDI,
																																				 this.mnuArchImpCostoI,
																																				 this.mnuRolesDistribucion,
																																				 this.mnuHerrContratos,
																																				 this.mnuHerrContratosIngreso,
																																				 this.mnuHerrContratosGeneraFact,
																																				 this.mnuFacLotTransfTrans,
																																				 this.mnuHerrContratosAdemdum,
																																				 this.mnuHerrContratosAlarma,
																																				 this.mnuHerrContratosAdemdums,
																																				 this.mnuFacLotCopiaArt,
																																				 this.mnuArchImpSeprofin,
																																				 this.btControlRemoto,
																																				 this.mnuFacLotFactReembol,
																																				 this.mnuContLotCuentaObligatoria,
																																				 this.mnuRolesAdicional,
																																				 this.MnuFactResCruzarFactReserva,
																																				 this.mnuArchImpMarceloVentas,
																																				 this.mnuRolesAdicionalFormulas,
																																				 this.MnuFactResOTrabajo,
																																				 this.MnuFactResConclusion,
																																				 this.MnuFactResOTrabajoServicio,
																																				 this.mnuHerrPresupC,
																																				 this.mnuHerrPresupCIngreso,
																																				 this.mnuHerrPresupCAnalisis,
																																				 this.MnuFactLotimp2,
																																				 this.mnuContTransConciliacion,
																																				 this.mnuHerrPresupCAvanceObra,
																																				 this.mnuHerrPresupCImportar,
																																				 this.MnuFactResJarComision,
																																				 this.lblMensaje,
																																				 this.barSkin,
																																				 this.mnuHerrSegGenGrupos,
																																				 this.mnuArchPropEmpresaR,
																																				 this.MnuEdicSegConfigUsuarios,
																																				 this.MnuFactContometros,
																																				 this.barButtonItem3,
																																				 this.barSubItem1,
																																				 this.mnuArchImpJuanMarcet,
																																				 this.barButtonItem4,
																																				 this.mnuArchImpFactRet,
																																				 this.mnuHerrAgenda,
																																				 this.mnuHerrPresupCSolicitudPago,
																																				 this.mnuHerrPresupCReportes,
																																				 this.mnuAyuJuego,
																																				 this.mnuHerrContratosAmortiza,
																																				 this.MnuFactSerie,
																																				 this.MnuFactPreFactura,
																																				 this.MnuDepTec,
																																				 this.MnuFacSerie,
																																				 this.barButtonItem5,
																																				 this.MnuHerrAlarma,
																																				 this.MnuFacSpif,
																																				 this.MnuFactCombos,
																																				 this.MnnuFactLotesPromociones,
																																				 this.barButtonItem6,
																																				 this.MnuCredito,
																																				 this.MnuCreditoSolRapida,
																																				 this.MnuFactLoteTransMult,
																																				 this.MnuFactLoteSpif,
																																				 this.mnuCrediPointSolicitud,
																																				 this.mnuCrediPointMantenedor,
																																				 this.mnuCrediPointVerTerreno,
																																				 this.mnuCrediPointGestiones,
																																				 this.mnuCrediPointGestVerCed,
																																				 this.barButtonItem11,
																																				 this.mnuCrediPointCatalogos,
																																				 this.mnuCrediPointGestRevCed,
																																				 this.MnuFacInv,
																																				 this.MnuFacOCImp,
																																				 this.MnuCreditoCotizador,
																																				 this.MnuFactLotesPromVen,
																																				 this.MnuContCuadreCaja,
																																				 this.MnuCreditoCuota,
																																				 this.MnuCreGastoCobran,
																																				 this.MnuCreCobranza,
																																				 this.MnuFactAnticipos,
																																				 this.mnuVentasComisiones,
																																				 this.MnuFacComConfiguracion,
																																				 this.MnuFactComCalificacion,
																																				 this.MnuFactDetalle,
																																				 this.MnuFactComConsulta,
																																				 this.MnuFacArtServicios,
																																				 this.barSubItem6,
																																				 this.barSubItem7,
																																				 this.MnuFactConsignacion,
																																				 this.MnuFacPagosAut,
																																				 this.MnuFactAutPagos,
																																				 this.MnuContAsigArt,
																																				 this.MnuFacSeriales,
																																				 this.FacGuiaRemision,
																																				 this.MnuContMantTarj,
																																				 this.MnuFactEstTransf,
																																				 this.mnuFactGuiaRemision,
																																				 this.mnuFactActPrecio,
																																				 this.MnuProdEstado,
																																				 this.mnuProduccionPendientes,
																																				 this.mnuCreLotesUnibanco,
																																				 this.MnuContDepositoVouchers,
																																				 this.MnuContCXP,
																																				 this.MnuContCXC,
																																				 this.mnuFactVentasConsignacion,
																																				 this.mnuDT,
																																				 this.MnuDTSTA,
																																				 this.MnuDTSolicitud,
																																				 this.barButtonItem16,
																																				 this.MnuCompras,
																																				 this.mnuVentas,
																																				 this.MnuBodega,
																																				 this.MnuContSRI,
																																				 this.mnuVentasResumenDiario,
																																				 this.mnuVentasCB,
																																				 this.mnuVentaLiqCons,
																																				 this.mnuVentaAutDistribuidor,
																																				 this.mnuVentasCotizacion,
																																				 this.mnuComprasComboZOOM,
																																				 this.mnuBodegaKardex,
																																				 this.mnuComprasLiquidaGastos,
																																				 this.mnuMatrizDsctos,
																																				 this.mnuComprasAgenda,
																																				 this.mnuCompraGestion,
																																				 this.mnuVentasComDetallePersonal,
																																				 this.mnuContResCXC,
																																				 this.mnuContResCXP,
																																				 this.mnuContRepAnt,
																																				 this.barButtonItem7,
																																				 this.mnuRolFiniquitos,
																																				 this.barButtonItem8,
																																				 this.mnuHerramientasNumeracion,
																																				 this.mnuContCCxCCP,
																																				 this.barButtonItem10,
																																				 this.barSubItem3,
																																				 this.mnuVentasReporteCompara,
																																				 this.mnuContabilidadAsDesc,
																																				 this.mnuContAntProv,
																																				 this.mnuDTIngresoMultiple,
																																				 this.mnuContRepCC,
																																				 this.mnuContASF,
																																				 this.mnuContMovBanc,
																																				 this.mnuContEstadoCheques,
																																				 this.mnuBodegaReportes,
																																				 this.mnuTomaInventario,
																																				 this.mnuBodegaCI,
																																				 this.mnuComprasRepSpifVend,
																																				 this.mnuContPrestamos,
																																				 this.mnuContTEB,
																																				 this.mnuVentasDA,
																																				 this.barButtonItem13,
																																				 this.barSubItem4,
																																				 this.mnuContReportes,
																																				 this.mnuContConsumosTC,
																																				 this.mnuBodegaDCI,
																																				 this.mnuContLCTC,
																																				 this.mnuProduccionTPT,
																																				 this.mnuContTCC,
																																				 this.barButtonItem14,
																																				 this.mnuDTRetiro,
																																				 this.mnuVentasImpresionDePrecios,
																																				 this.barButtonItem9,
																																				 this.mnuHerrSegCD,
																																				 this.mnuComprasReportes,
																																				 this.mnuCmpRepVM,
																																				 this.mnuRepCmpTransf,
																																				 this.barButtonItem12,
																																				 this.mnuBodegaCierreInventario,
																																				 this.barButtonItem17,
																																				 this.barButtonItem18,
																																				 this.mnuGPPD,
																																				 this.barButtonItem19,
																																				 this.mnuCPReportes,
																																				 this.mnuCPReportesCliPoint,
																																				 this.mnuCPInformeCobrador,
																																				 this.mnuCompraReportesCPA,
																																				 this.mnuVentasReportesPorMarca,
																																				 this.mnuHerramientasCRIVA,
																																				 this.mnuComprasPromociones,
																																				 this.mnuVentaPromocionesActivas,
																																				 this.mnuBodegaIF,
																																				 this.mnuVentasRVML,
																																				 this.mnuCrediPointDFG,
																																				 this.mnuCrediPointMEF,
																																				 this.mnuCPCP,
																																				 this.mnuCPIO,
																																				 this.mnuBodegaSectorizacion,
																																				 this.mnuComprasTEL,
																																				 this.mnuVentasIPP,
																																				 this.mnuContBancosCBA,
																																				 this.mnuBodegaInvVal,
																																				 this.mnuCPLCS,
																																				 this.mnuCPRepEquifax,
																																				 this.mnuCPMFCP,
																																				 this.mnuVentasRepMU,
																																				 this.mnuContRepGastos,
																																				 this.mnuCPRepCP,
																																				 this.barButtonItem20,
																																				 this.mnuContRepFacElec,
																																				 this.barToolbarsListItem1,
																																				 this.barButtonItem21,
																																				 this.barSubItem8,
																																				 this.mnuLotte,
																																				 this.mnuHabitaciones,
																																				 this.mnuTmaHabitacion,
																																				 this.mnuTmpEstadia,
																																				 this.mnuTmpHabitaciones,
																																				 this.mnuCuadroHabitaciones,
																																				 this.mnuCobrPrenda,
																																				 this.mnuKardexLot,
																																				 this.mnuPromociones,
																																				 this.mnuReservas,
																																				 this.mnuCpLbp,
																																				 this.mnuVentasReportesAIM,
																																				 this.mnuVentasReportesRF,
																																				 this.barButtonItem22,
																																				 this.barButtonItem23,
																																				 this.barSubItem9,
																																				 this.barButtonItem24,
																																				 this.barButtonItem25,
																																				 this.barButtonItem26,
																																				 this.mnuVentasVentasCP,
																																				 this.barButtonItem27,
																																				 this.mnuHerrALF,
																																				 this.mnuBodegaIngSeriales,
																																				 this.barButtonItem28,
																																				 this.barButtonItem29,
																																				 this.barButtonItem30,
																																				 this.mnuBodegaGRC,
																																				 this.barButtonItem31,
																																				 this.barButtonItem32,
																																				 this.mnuCPRRCS,
																																				 this.mnuComprasGDA,
																																				 this.mnuLoteNotasDeEntrega,
																																				 this.mnuCPRD,
																																				 this.mnuCPRepRDD,
																																				 this.mnuCPVEP,
																																				 this.mnuVentasARP,
																																				 this.mnuCPRepECP,
																																				 this.mnuCPEDC,
																																				 this.mnuComprasRepFP,
																																				 this.barButtonItem33,
																																				 this.barButtonItem34,
																																				 this.mnuRolesCom_AV,
																																				 this.barButtonItem36,
																																				 this.mnuCPHRC,
																																				 this.mnuVentaRMU,
																																				 this.mnuCPRRC,
																																				 this.barButtonItem37,
																																				 this.barButtonItem38,
																																				 this.mnuDTIRM,
																																				 this.mnuDTIRMP,
																																				 this.mnuRolesCom_PA,
																																				 this.mnuRolesCom_PL,
																																				 this.mnuRolesCom_PML,
																																				 this.mnuRolesComisiones,
																																				 this.mnuVentasPC,
																																				 this.barButtonItem35,
																																				 this.mnuRolPA,
																																				 this.barButtonItem39,
																																				 this.mnuCPCobranzas,
																																				 this.mnuCPCobranzasRCP,
																																				 this.mnuVentasPlanifVentas,
																																				 this.mnuRolComisionesPC,
																																				 this.barButtonItem40,
																																				 this.mnuVentasResumenComisiones,
																																				 this.mnuCPReportesRCV,
																																				 this.mnuRolesPagoComisiones,
																																				 this.mnuRolesComAJV,
																																				 this.mnuCPDFGP,
																																				 this.mnuComprasAPC,
																																				 this.barButtonItem41,
																																				 this.mnuRolesReportes,
																																				 this.mnuRolesReportesPVF,
																																				 this.barButtonItem42,
																																				 this.mnuCom_PAGE,
																																				 this.mnuCom_PMLGE,
																																				 this.mnuCom_PCGE,
																																				 this.mnuVentasRepGE,
																																				 this.mnuCom_Reportes,
																																				 this.mnuCom_ReportesRM,
																																				 this.mnuCobranzasRGC,
																																				 this.mnuRolesReportesVXM,
																																				 this.mnuVentasMarcador,
																																				 this.barButtonItem44,
																																				 this.mnuReportesBodega,
																																				 this.barButtonItem45,
																																				 this.mnuCPCartera,
																																				 this.barButtonItem46,
																																				 this.barButtonItem47,
																																				 this.barSubItem10,
																																				 this.barButtonItem48,
																																				 this.barSubItem11,
																																				 this.mnuCPLotesPPagos,
																																				 this.barButtonItem49,
																																				 this.mnuVentasReportesComparativos,
																																				 this.mnuCPReportesGDCM,
																																				 this.mnuBodegaIPR,
																																				 this.barButtonItem50,
																																				 this.mnuDTIRAT,
																																				 this.mnuDTIREP,
																																				 this.mnuDTIRRPT,
																																				 this.mnuDTIREPC,
																																				 this.mnuContRepCajasChicas,
																																				 this.barButtonItem55,
																																				 this.mnuCPReportesAbonos,
																																				 this.mnuVentasReportesVentasPorMes,
																																				 this.mnuCPReportesNDC,
																																				 this.mnuContReportesGA,
																																				 this.barButtonItem56,
																																				 this.mnuContBancosCP,
																																				 this.mnuVentasReportesTC,
																																				 this.mnuRolesReportesUniformes,
																																				 this.barButtonItem57,
																																				 this.mnuContBancosDPL,
																																				 this.mnuContBancosCDPL,
																																				 this.mnuDTIRPPL,
																																				 this.mnuDTIRAPD,
																																				 this.mnuDTIRSNC,
																																				 this.barButtonItem61,
																																				 this.mnuDTIRREP,
																																				 this.mnuDTIRRMA,
																																				 this.mnuDTIRRRP,
																																				 this.mnuDTIRRTE,
																																				 this.mnuDTIRRPG,
																																				 this.mnuDTIRRSR,
																																				 this.barButtonItem66,
																																				 this.mnuVentasRepRVF,
																																				 this.barButtonItem67,
																																				 this.mnuDTProductosDeFeria,
																																				 this.mnuVentasPlanifJefeVentas,
																																				 this.barButtonItem69,
																																				 this.barButtonItem68,
																																				 this.mnuRolesReportesNominaGeneral,
																																				 this.barButtonItem70,
																																				 this.barButtonItem71,
																																				 this.mnuHerramientasPTP,
																																				 this.mnuHerramientasAP,
																																				 this.mnuRolesCom_EPV,
																																				 this.mnuRolPersonalPrueba,
																																				 this.mnuRolComPIA,
																																				 this.mnuVRE,
																																				 this.mnuVentasReportesDTC,
																																				 this.mnuVentasReportesEncuestas,
																																				 this.barButtonItem74,
																																				 this.barSubItem14,
																																				 this.mnuLoteRepProcSIM,
																																				 this.mnuCompraCSA,
																																				 this.barButtonItem75,
																																				 this.mnuVentasGCCC,
																																				 this.barButtonItem76,
																																				 this.mnuBodegaRM,
																																				 this.mnuCPReportesBR,
																																				 this.mnuBodegaRepKS,
																																				 this.barButtonItem72,
																																				 this.mnuDTIRFDT,
																																				 this.mnuDTIRSIN,
																																				 this.mnuDTIRCDE,
																																				 this.barButtonItem51,
																																				 this.mnuRolCom_PorcCarpa,
																																				 this.mnuReporteVentasMensualesPorLocal,
																																				 this.mnuVentasIPR,
																																				 this.mnuBodegaCDI,
																																				 this.frmDTTransfLocales,
																																				 this.mnuDTRA,
																																				 this.mnuDTReporteEstadoIngresos,
																																				 this.mnuCPCobranzasLPCA,
																																				 this.mnuVentasReportesVA,
																																				 this.mnuDTAprobacionRemates,
																																				 this.mnuDTASSF,
																																				 this.mnuContabilidadProvisiones,
																																				 this.mnuContabilidadFlujoPMP,
																																				 this.mnuContFPPP,
																																				 this.mnuCPReportesREC,
																																				 this.barButtonItem58,
																																				 this.mnuMetricas,
																																				 this.barButtonItem59,
																																				 this.mnuMetricasRRHH,
																																				 this.barButtonItem60,
																																				 this.barButtonItem62,
																																				 this.mnuBodegaReportesECDI,
																																				 this.barButtonItem63,
																																				 this.mnuPCGEx,
																																				 this.mnuSTA,
																																				 this.mnuVentasReportesFFP,
																																				 this.mnuVentasPFR,
																																				 this.mnuVentasEOC,
																																				 this.mnuLottePlanificacion,
																																				 this.mnuVentasReportesVPD,
																																				 this.mnuVentasGCR,
																																				 this.mnuLotteReportesIT,
																																				 this.barButtonItem73,
																																				 this.mnuLotteCompetencia,
																																				 this.mnuLotteReporteIPH,
																																				 this.mnuRolReportesAE,
																																				 this.mnuCrediPOINTRMR,
																																				 this.mnuCrediPOINTReportesRMR,
																																				 this.mnuVentasReportesCXC,
																																				 this.mnuCPReportesEC,
																																				 this.mnuPSRDC,
																																				 this.mnuRolesReportesHHT,
																																				 this.mnuCPCobranzasCCPPC,
																																				 this.mnuVentasReportesVPA,
																																				 this.mnuVentasRP,
																																				 this.mnuVentasECW,
																																				 this.mnuHerrVerifDepositos,
																																				 this.mnuHerrSegClientes,
																																				 this.mnuHerrDepositos,
																																				 this.mnuContabilidadFPPP,
																																				 this.mnuVentasReportesVentasPorArticulos,
																																				 this.mnuVentasReportesVentasarticuloPorMes,
																																				 this.mnuComprasRepCodigoArticulos,
																																				 this.mnuDTSubGrupo,
																																				 this.mnuCPCobranzasEDC,
																																				 this.barButtonItem52,
																																				 this.mnuComprasCatalogo,
																																				 this.mnuBodegaRepMD,
																																				 this.mnuContRB,
																																				 this.mnuBodegaTIM,
																																				 this.mnuBodegaTRM,
																																				 this.mnuVentasRPA,
																																				 this.mnuRolesComisionesPS,
																																				 this.mnuRolesComisionesPCS,
																																				 this.barButtonItem53,
																																				 this.mnuTarjetasDeIngreso,
																																				 this.mnuVentasCNOS,
																																				 this.mnuDirectorio,
																																				 this.mnuCPReportesRRM,
																																				 this.mnuComprasProformas,
																																				 this.mnuRolReportesIND,
																																				 this.mnuCPReportesCVPM,
																																				 this.mnuCPCobranzasMRPR,
																																				 this.mnuHerramientasPTA,
																																				 this.mnuCPCobranzasGDC,
																																				 this.mnuVentaRMUA,
																																				 this.mnuContRepEstadoCB,
																																				 this.mnuContCaja,
																																				 this.mnuContEstadosCajas,
																																				 this.mnuContIngresoDiarioCaja,
																																				 this.mnuContIngresosDeCajas,
																																				 this.mnuContEstadoGeneralDeCaja,
																																				 this.mnuContRepPresupAlm,
																																				 this.mnuHerramientasAPL,
																																				 this.mnuHerramientasASDP,
																																				 this.mnuCPRevision,
																																				 this.barButtonItem43,
																																				 this.mnuCPCobranzasCarteraAIP,
																																				 this.mnuCPCobranzasRPE,
																																				 this.mnuRolesDecimoTercerSueldo,
																																				 this.barButtonItem54,
																																				 this.mnuCPCRD,
																																				 this.mnuCPCDRDC,
																																				 this.mnuCPCDRVC,
																																				 this.mnuCPCDSR,
																																				 this.mnuCPCDRSLT,
																																				 this.mnuCPCDLCPG,
																																				 this.mnuBodegaReportesKardexCosto,
																																				 this.barButtonItem64,
																																				 this.barLinkContainerItem1,
																																				 this.barSubItem5,
																																				 this.barButtonItem65,
																																				 this.mnuRPtBodega,
																																				 this.mnuRPtContabilidad,
																																				 this.barButtonItem77,
																																				 this.barButtonItem78,
																																				 this.mnuRPtRoles,
																																				 this.barButtonItem79,
																																				 this.mnuRPtCompras,
																																				 this.barButtonItem80,
																																				 this.mnuRPtVentas,
																																				 this.mnuRPtCREPOINT,
																																				 this.mnuRPtDT,
																																				 this.mnuRPtPeonyServicios,
																																				 this.mnuRPtBodegaInvFEcha,
																																				 this.mnuRPtBodegaRptTrans,
																																				 this.mnuRPtBodegaInvVal,
																																				 this.mnuRPtBodegaKXSer,
																																				 this.mnuRPtBodegaEsConInv,
																																				 this.mnuRPtBodegaMovDesc,
																																				 this.mnuRPtContabilidadCont,
																																				 this.mnuRPtContabilidadRet,
																																				 this.mnuRPtContabilidadAsDEsc,
																																				 this.mnuRPtContabilidadGast,
																																				 this.mnuRPtContabilidadFactEl,
																																				 this.mnuRPtContabilidadCaCh,
																																				 this.mnuRPtContabilidadGaAn,
																																				 this.mnuRPtContabilidadEstaConB,
																																				 this.mnuRPtContabilidadPrXAlm,
																																				 this.mnuRPtRolesPFA,
																																				 this.mnuRPtRolesVeXM,
																																				 this.mnuRPtRolesUni,
																																				 this.mnuRPtRolesNG,
																																				 this.mnuRPtRolesAEM,
																																				 this.mnuRPtRolesHHT,
																																				 this.mnuRPtRolesIXD,
																																				 this.mnuRPtComprasCPA,
																																				 this.mnuRPtComprasT,
																																				 this.mnuRPtComprasRAC,
																																				 this.mnuRPtComprasCA,
																																				 this.mnuIngreso,
																																				 this.mnuRPTVentasPorcMargUtilidad,
																																				 this.mnuRPtBodegaKCC,
																																				 this.mnuRPTVentasPorcMarca,
																																				 this.mnuRPTVentasPorcMarcaUtilid,
																																				 this.mnuRPTVentasMargUti,
																																				 this.mnuRPTVentasReMarUti,
																																				 this.mnuRPTVentasMaUXArt,
																																				 this.mnuRPTVentasAnInMEn,
																																				 this.mnuRPTVentasRevFact,
																																				 this.mnuRPTVentasXDia,
																																				 this.mnuRPTVentasNFactPForPago,
																																				 this.mnuRPTVentasGaraExt,
																																				 this.mnuRPTVentasXFecha,
																																				 this.mnuRPTVentasComDeVentas,
																																				 this.mnuRPTVentasANuXLo,
																																				 this.mnuRPTVentasMenXLo,
																																				 this.mnuRPTVentasTarCredi,
																																				 this.mnuRPTVentasDetTc,
																																				 this.mnuRPTVentasEncuestas,
																																				 this.mnuRPTVentasReCXCo,
																																				 this.mnuRPTVentasXano,
																																				 this.mnuRPTVentasXArtic,
																																				 this.mnuRPTVentasArtXMes,
																																				 this.barButtonItem81,
																																				 this.barSubItem12,
																																				 this.mnuRPTVentasComisionesConf,
																																				 this.mnuRPTVentasComisionesCali,
																																				 this.mnuRPTVentasComisionesDetGe,
																																				 this.mnuRPTVentasComisionesConsulta,
																																				 this.mnuCompraAODC,
																																				 this.barButtonItem83,
																																				 this.mnuCPCobCM,
																																				 this.mnuReporVentasPDGA,
																																				 this.mnuCRResumenCredito,
																																				 this.barButtonItem82,
																																				 this.mnuCom_ReporteLM,
																																				 this.mnuBodegaNI,
																																				 this.mnuComprasLAC,
																																				 this.mnuCPComisiones,
																																				 this.mnuCPComisionesAsigCobr,
																																				 this.mnuCPComisionesTrmOpe,
																																				 this.mnuCPCobranzasCboOpe,
																																				 this.mnuCPComisionesCom,
																																				 this.mnuCreditoGesCall,
																																				 this.mnuCreditoProReal,
																																				 this.mnuRPPagoAgil,
																																				 this.mnuCPCobranzasCERTIFICADO,
																																				 this.mnuRPtComprasPuC,
																																				 this.barButtonItem84,
																																				 this.mnuCPANDC,
																																				 this.mnuComprasPromoFact,
																																				 this.barButtonItem85,
																																				 this.mnuSolNC,
																																				 this.mnuRolPorCajero,
																																				 this.barButtonItem86,
																																				 this.barButtonItem87,
																																				 this.barButtonItem88,
																																				 this.mnuRPtCREPOINTFDG,
																																				 this.mnuRPtCREPOINRMD,
																																				 this.mnuCPCobranzasRPP,
																																				 this.mnuReporVentasDFI,
																																				 this.mnuRPtContabilidadSCosecha});
			this.lblTiempos.MainMenu = this.bar1;
			this.lblTiempos.MaxItemId = 1003;
			this.lblTiempos.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																				this.repositoryItemTextEdit1,
																																																				this.repositoryItemComboBox1});
			// 
			// bar1
			// 
			this.bar1.BarName = "Custom 1";
			this.bar1.DockCol = 0;
			this.bar1.DockRow = 0;
			this.bar1.DockStyle = DevExpress.XtraBars.BarDockStyle.Top;
			this.bar1.FloatLocation = new System.Drawing.Point(51, 131);
			this.bar1.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																											new DevExpress.XtraBars.LinkPersistInfo(this.MnuArch),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.MnuEdic),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.MnuBodega),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.MnuHerr_),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.MnuCont),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.MnuRol),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.MnuCompras),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentas),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.MnuCredito),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.MnuProduc),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.mnuDT),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.mnuLotte),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.MnuVentana),
																																											new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.MnuFact, false),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.mnuMetricas),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.mnuPCGEx),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.mnuSTA),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.mnuDirectorio),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.barSubItem5),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.mnuIngreso),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem86),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem87)});
			this.bar1.OptionsBar.MultiLine = true;
			this.bar1.OptionsBar.UseWholeRow = true;
			this.bar1.Text = "Menu";
			// 
			// MnuArch
			// 
			this.MnuArch.Caption = "&Archivo";
			this.MnuArch.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArch.Id = 0;
			this.MnuArch.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchNuevo),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchAbrir),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchGuardar, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchImpoExpo),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchProp, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchConfImp),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem14, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem12),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem17),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchSalir, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem53)});
			this.MnuArch.Name = "MnuArch";
			// 
			// MnuArchNuevo
			// 
			this.MnuArchNuevo.Caption = "Nuevo";
			this.MnuArchNuevo.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchNuevo.Id = 3;
			this.MnuArchNuevo.ImageIndex = 8;
			this.MnuArchNuevo.Name = "MnuArchNuevo";
			this.MnuArchNuevo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchNuevo_ItemClick);
			// 
			// MnuArchAbrir
			// 
			this.MnuArchAbrir.Caption = "Abrir";
			this.MnuArchAbrir.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchAbrir.Id = 5;
			this.MnuArchAbrir.ImageIndex = 4;
			this.MnuArchAbrir.Name = "MnuArchAbrir";
			this.MnuArchAbrir.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchIniSes_ItemClick);
			// 
			// MnuArchGuardar
			// 
			this.MnuArchGuardar.Caption = "&Guardar";
			this.MnuArchGuardar.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchGuardar.Id = 13;
			this.MnuArchGuardar.ImageIndex = 9;
			this.MnuArchGuardar.Name = "MnuArchGuardar";
			this.MnuArchGuardar.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchGuardar_ItemClick);
			// 
			// MnuArchImpoExpo
			// 
			this.MnuArchImpoExpo.Caption = "Importar";
			this.MnuArchImpoExpo.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchImpoExpo.Id = 23;
			this.MnuArchImpoExpo.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchImpoLat),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchImpoOtros),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuArchImpOmniLife, true),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuArchImpTelcosan),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mneArchImpInterActive, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mneArchImpYellow, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mneArchImpClaseEcuador, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuArchImpFresh, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuArchImpCash, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuArchImpQBI, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuArchImpCashE, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuArchImpCerv, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuArchImpResCDI, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuArchImpCostoI, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuArchImpSeprofin, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.barSubItem1, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuArchImpJuanMarcet, false),
																																																 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuArchImpFactRet, false)});
			this.MnuArchImpoExpo.Name = "MnuArchImpoExpo";
			// 
			// MnuArchImpoLat
			// 
			this.MnuArchImpoLat.Caption = "Importar de Latinium";
			this.MnuArchImpoLat.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchImpoLat.Id = 24;
			this.MnuArchImpoLat.Name = "MnuArchImpoLat";
			this.MnuArchImpoLat.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArch_Impo_CIAD_ItemClick);
			// 
			// MnuArchImpoOtros
			// 
			this.MnuArchImpoOtros.Caption = "Importar de otros programas";
			this.MnuArchImpoOtros.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchImpoOtros.Id = 25;
			this.MnuArchImpoOtros.Name = "MnuArchImpoOtros";
			this.MnuArchImpoOtros.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArch_Impo_Otros_ItemClick);
			// 
			// mnuArchImpOmniLife
			// 
			this.mnuArchImpOmniLife.Caption = "OmniLife";
			this.mnuArchImpOmniLife.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchImpOmniLife.Id = 265;
			this.mnuArchImpOmniLife.Name = "mnuArchImpOmniLife";
			this.mnuArchImpOmniLife.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpOmniLife.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpOmniLife_ItemClick);
			// 
			// mnuArchImpTelcosan
			// 
			this.mnuArchImpTelcosan.Caption = "SuperMas";
			this.mnuArchImpTelcosan.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchImpTelcosan.Id = 278;
			this.mnuArchImpTelcosan.Name = "mnuArchImpTelcosan";
			this.mnuArchImpTelcosan.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpTelcosan.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpTelcosan_ItemClick);
			// 
			// mneArchImpInterActive
			// 
			this.mneArchImpInterActive.Caption = "InterActive";
			this.mneArchImpInterActive.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mneArchImpInterActive.Id = 282;
			this.mneArchImpInterActive.Name = "mneArchImpInterActive";
			this.mneArchImpInterActive.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mneArchImpInterActive.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mneArchImpInterActive_ItemClick);
			// 
			// mneArchImpYellow
			// 
			this.mneArchImpYellow.Caption = "Yellow Pepper";
			this.mneArchImpYellow.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mneArchImpYellow.Id = 287;
			this.mneArchImpYellow.Name = "mneArchImpYellow";
			this.mneArchImpYellow.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mneArchImpYellow.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mneArchImpFuneraria_ItemClick);
			// 
			// mneArchImpClaseEcuador
			// 
			this.mneArchImpClaseEcuador.Caption = "Clase Ecuador";
			this.mneArchImpClaseEcuador.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mneArchImpClaseEcuador.Id = 290;
			this.mneArchImpClaseEcuador.Name = "mneArchImpClaseEcuador";
			this.mneArchImpClaseEcuador.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mneArchImpClaseEcuador.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mneArchImpClaseEcuador_ItemClick);
			// 
			// mnuArchImpFresh
			// 
			this.mnuArchImpFresh.Caption = "Fresh Logistics";
			this.mnuArchImpFresh.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchImpFresh.Id = 294;
			this.mnuArchImpFresh.Name = "mnuArchImpFresh";
			this.mnuArchImpFresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpFresh.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpFresh_ItemClick);
			// 
			// mnuArchImpCash
			// 
			this.mnuArchImpCash.Caption = "Cash Management";
			this.mnuArchImpCash.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchImpCash.Id = 295;
			this.mnuArchImpCash.Name = "mnuArchImpCash";
			this.mnuArchImpCash.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpCash.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpCash_ItemClick);
			// 
			// mnuArchImpQBI
			// 
			this.mnuArchImpQBI.Caption = "Que buena idea";
			this.mnuArchImpQBI.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchImpQBI.Id = 382;
			this.mnuArchImpQBI.Name = "mnuArchImpQBI";
			this.mnuArchImpQBI.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpQBI.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpQBI_ItemClick);
			// 
			// mnuArchImpCashE
			// 
			this.mnuArchImpCashE.Caption = "Cash Management Envio";
			this.mnuArchImpCashE.Id = 400;
			this.mnuArchImpCashE.Name = "mnuArchImpCashE";
			this.mnuArchImpCashE.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpCashE.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpCashE_ItemClick);
			// 
			// mnuArchImpCerv
			// 
			this.mnuArchImpCerv.Caption = "Cerveceria";
			this.mnuArchImpCerv.Id = 401;
			this.mnuArchImpCerv.Name = "mnuArchImpCerv";
			this.mnuArchImpCerv.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpCerv.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpCerv_ItemClick_1);
			// 
			// mnuArchImpResCDI
			// 
			this.mnuArchImpResCDI.Caption = "Resumen Costos D - I";
			this.mnuArchImpResCDI.Id = 403;
			this.mnuArchImpResCDI.Name = "mnuArchImpResCDI";
			this.mnuArchImpResCDI.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpResCDI.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpResCDI_ItemClick);
			// 
			// mnuArchImpCostoI
			// 
			this.mnuArchImpCostoI.Caption = "Costos Indirectos Rio Verde";
			this.mnuArchImpCostoI.Id = 404;
			this.mnuArchImpCostoI.Name = "mnuArchImpCostoI";
			this.mnuArchImpCostoI.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpCostoI.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpCostoI_ItemClick);
			// 
			// mnuArchImpSeprofin
			// 
			this.mnuArchImpSeprofin.Caption = "Seprofin";
			this.mnuArchImpSeprofin.Id = 416;
			this.mnuArchImpSeprofin.Name = "mnuArchImpSeprofin";
			this.mnuArchImpSeprofin.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpSeprofin.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btSeprofin_ItemClick);
			// 
			// barSubItem1
			// 
			this.barSubItem1.Caption = "Gruas";
			this.barSubItem1.Id = 445;
			this.barSubItem1.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuArchImpGuiaGruas),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuArchImpGruasCompra),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuArchImpGuiaVenta)});
			this.barSubItem1.Name = "barSubItem1";
			this.barSubItem1.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuArchImpGuiaGruas
			// 
			this.mnuArchImpGuiaGruas.Caption = "Guia Remision Gruas";
			this.mnuArchImpGuiaGruas.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchImpGuiaGruas.Id = 300;
			this.mnuArchImpGuiaGruas.Name = "mnuArchImpGuiaGruas";
			this.mnuArchImpGuiaGruas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpGuiaGruas_ItemClick);
			// 
			// mnuArchImpGruasCompra
			// 
			this.mnuArchImpGruasCompra.Caption = "Guia Remision Ventas";
			this.mnuArchImpGruasCompra.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchImpGruasCompra.Id = 301;
			this.mnuArchImpGruasCompra.Name = "mnuArchImpGruasCompra";
			this.mnuArchImpGruasCompra.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpGruasCompra_ItemClick);
			// 
			// mnuArchImpGuiaVenta
			// 
			this.mnuArchImpGuiaVenta.Caption = "Guia Remision Compras";
			this.mnuArchImpGuiaVenta.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchImpGuiaVenta.Id = 302;
			this.mnuArchImpGuiaVenta.Name = "mnuArchImpGuiaVenta";
			this.mnuArchImpGuiaVenta.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpGuiaVenta_ItemClick);
			// 
			// mnuArchImpJuanMarcet
			// 
			this.mnuArchImpJuanMarcet.Caption = "Importacion Personalizada";
			this.mnuArchImpJuanMarcet.Id = 446;
			this.mnuArchImpJuanMarcet.Name = "mnuArchImpJuanMarcet";
			this.mnuArchImpJuanMarcet.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpJuanMarcet.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpJuanMarcet_ItemClick);
			// 
			// mnuArchImpFactRet
			// 
			this.mnuArchImpFactRet.Caption = "Factura - Retencion";
			this.mnuArchImpFactRet.Id = 448;
			this.mnuArchImpFactRet.Name = "mnuArchImpFactRet";
			this.mnuArchImpFactRet.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchImpFactRet.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchImpFactRet_ItemClick);
			// 
			// MnuArchProp
			// 
			this.MnuArchProp.Caption = "Propiedades";
			this.MnuArchProp.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchProp.Id = 33;
			this.MnuArchProp.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuArchPropGen),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchPropCont, true),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchPropFact),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchPropNumE),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.MnuArchPropRol),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuArchPropProd),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuArchPropEmpresa, true),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuArchPropEmpresaR),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuArchPropConfig)});
			this.MnuArchProp.Name = "MnuArchProp";
			// 
			// mnuArchPropGen
			// 
			this.mnuArchPropGen.Caption = "Generales";
			this.mnuArchPropGen.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchPropGen.Id = 214;
			this.mnuArchPropGen.Name = "mnuArchPropGen";
			this.mnuArchPropGen.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchPropGen_ItemClick);
			// 
			// MnuArchPropCont
			// 
			this.MnuArchPropCont.Caption = "Contabilidad";
			this.MnuArchPropCont.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchPropCont.Id = 34;
			this.MnuArchPropCont.Name = "MnuArchPropCont";
			this.MnuArchPropCont.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchProp_Cont_ItemClick);
			// 
			// MnuArchPropFact
			// 
			this.MnuArchPropFact.Caption = "Facturación";
			this.MnuArchPropFact.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchPropFact.Id = 35;
			this.MnuArchPropFact.Name = "MnuArchPropFact";
			this.MnuArchPropFact.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchPropFact_ItemClick);
			// 
			// MnuArchPropNumE
			// 
			this.MnuArchPropNumE.Caption = "Numeración Estación";
			this.MnuArchPropNumE.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchPropNumE.Id = 297;
			this.MnuArchPropNumE.Name = "MnuArchPropNumE";
			this.MnuArchPropNumE.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.MnuArchPropNumE.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchPropNumE_ItemClick);
			// 
			// MnuArchPropRol
			// 
			this.MnuArchPropRol.Caption = "Rol de Pagos";
			this.MnuArchPropRol.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchPropRol.Id = 36;
			this.MnuArchPropRol.Name = "MnuArchPropRol";
			this.MnuArchPropRol.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchProp_Rol_ItemClick);
			// 
			// mnuArchPropProd
			// 
			this.mnuArchPropProd.Caption = "Producción";
			this.mnuArchPropProd.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchPropProd.Id = 219;
			this.mnuArchPropProd.Name = "mnuArchPropProd";
			this.mnuArchPropProd.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuArchPropProd.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchPropProd_ItemClick);
			// 
			// mnuArchPropEmpresa
			// 
			this.mnuArchPropEmpresa.Caption = "Empresas";
			this.mnuArchPropEmpresa.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchPropEmpresa.Id = 213;
			this.mnuArchPropEmpresa.Name = "mnuArchPropEmpresa";
			this.mnuArchPropEmpresa.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchPropEmpresa_ItemClick);
			// 
			// mnuArchPropEmpresaR
			// 
			this.mnuArchPropEmpresaR.Caption = "Empresas Relacionadas";
			this.mnuArchPropEmpresaR.Id = 441;
			this.mnuArchPropEmpresaR.Name = "mnuArchPropEmpresaR";
			this.mnuArchPropEmpresaR.TextAlignment = System.Drawing.StringAlignment.Near;
			this.mnuArchPropEmpresaR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuArchPropEmpresaR_ItemClick);
			// 
			// mnuArchPropConfig
			// 
			this.mnuArchPropConfig.Caption = "Configuración";
			this.mnuArchPropConfig.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mnuArchPropConfig.Id = 393;
			this.mnuArchPropConfig.Name = "mnuArchPropConfig";
			this.mnuArchPropConfig.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btConfig_ItemClick);
			// 
			// MnuArchConfImp
			// 
			this.MnuArchConfImp.Caption = "Configurar Impresión";
			this.MnuArchConfImp.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchConfImp.Id = 38;
			this.MnuArchConfImp.ImageIndex = 6;
			this.MnuArchConfImp.Name = "MnuArchConfImp";
			this.MnuArchConfImp.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchConfImp_ItemClick);
			// 
			// barButtonItem14
			// 
			this.barButtonItem14.Caption = "CRM Articulos Clientes";
			this.barButtonItem14.Id = 582;
			this.barButtonItem14.Name = "barButtonItem14";
			this.barButtonItem14.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.barButtonItem14.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem14_ItemClick);
			// 
			// barButtonItem12
			// 
			this.barButtonItem12.Caption = "Ajusta CXC CXP";
			this.barButtonItem12.Id = 590;
			this.barButtonItem12.Name = "barButtonItem12";
			this.barButtonItem12.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.barButtonItem12.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem12_ItemClick);
			// 
			// barButtonItem17
			// 
			this.barButtonItem17.Caption = "Facturas EM";
			this.barButtonItem17.Id = 592;
			this.barButtonItem17.Name = "barButtonItem17";
			this.barButtonItem17.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.barButtonItem17.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem17_ItemClick);
			// 
			// MnuArchSalir
			// 
			this.MnuArchSalir.Caption = "Salir            ";
			this.MnuArchSalir.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.MnuArchSalir.Id = 39;
			this.MnuArchSalir.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.F4));
			this.MnuArchSalir.Name = "MnuArchSalir";
			this.MnuArchSalir.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchSalir_ItemClick);
			// 
			// barButtonItem53
			// 
			this.barButtonItem53.Caption = "Resumen De Comisiones Semanal";
			this.barButtonItem53.Id = 860;
			this.barButtonItem53.Name = "barButtonItem53";
			this.barButtonItem53.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem53_ItemClick_2);
			// 
			// MnuEdic
			// 
			this.MnuEdic.Caption = "&Edición";
			this.MnuEdic.CategoryGuid = new System.Guid("d9189416-8661-46d9-99d3-d809add0384f");
			this.MnuEdic.Id = 1;
			this.MnuEdic.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuEdicDesh),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuEdicReha),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuEdicCortar, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuEdicCopi),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuEdicPega)});
			this.MnuEdic.MergeOrder = 2;
			this.MnuEdic.Name = "MnuEdic";
			this.MnuEdic.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// MnuEdicDesh
			// 
			this.MnuEdicDesh.Caption = "Deshacer         ";
			this.MnuEdicDesh.CategoryGuid = new System.Guid("d9189416-8661-46d9-99d3-d809add0384f");
			this.MnuEdicDesh.Id = 41;
			this.MnuEdicDesh.ImageIndex = 10;
			this.MnuEdicDesh.Name = "MnuEdicDesh";
			this.MnuEdicDesh.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuEdicDesh_ItemClick);
			// 
			// MnuEdicReha
			// 
			this.MnuEdicReha.Caption = "Rehacer  ";
			this.MnuEdicReha.CategoryGuid = new System.Guid("d9189416-8661-46d9-99d3-d809add0384f");
			this.MnuEdicReha.Id = 42;
			this.MnuEdicReha.ImageIndex = 7;
			this.MnuEdicReha.Name = "MnuEdicReha";
			// 
			// MnuEdicCortar
			// 
			this.MnuEdicCortar.Caption = "Cortar       Ctrl + X";
			this.MnuEdicCortar.CategoryGuid = new System.Guid("d9189416-8661-46d9-99d3-d809add0384f");
			this.MnuEdicCortar.Id = 43;
			this.MnuEdicCortar.ImageIndex = 11;
			this.MnuEdicCortar.Name = "MnuEdicCortar";
			this.MnuEdicCortar.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuEdicCortar_ItemClick);
			// 
			// MnuEdicCopi
			// 
			this.MnuEdicCopi.Caption = "Copiar       Ctrl + C";
			this.MnuEdicCopi.CategoryGuid = new System.Guid("d9189416-8661-46d9-99d3-d809add0384f");
			this.MnuEdicCopi.Id = 44;
			this.MnuEdicCopi.ImageIndex = 0;
			this.MnuEdicCopi.Name = "MnuEdicCopi";
			this.MnuEdicCopi.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuEdicCopi_ItemClick);
			// 
			// MnuEdicPega
			// 
			this.MnuEdicPega.Caption = "Pegar        Ctrl + V";
			this.MnuEdicPega.CategoryGuid = new System.Guid("d9189416-8661-46d9-99d3-d809add0384f");
			this.MnuEdicPega.Id = 45;
			this.MnuEdicPega.ImageIndex = 5;
			this.MnuEdicPega.Name = "MnuEdicPega";
			// 
			// MnuBodega
			// 
			this.MnuBodega.Caption = "&Bodega";
			this.MnuBodega.Id = 528;
			this.MnuBodega.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaTRM),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactSerie),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFacInv, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaKardex),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactLoteSpif),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFacSeriales, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaIngSeriales),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaIPR),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuFactTransf, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactLoteTransMult),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuComprasTEL),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactEstTransf),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuFactGuiaRemision, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaGRC),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactEnvTransferencia, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactRecTransferencia),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaNI),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaCDI, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaRM),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaCI, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaDCI),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaCierreInventario),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuReportesBodega, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuFactBodegas, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaSectorizacion)});
			this.MnuBodega.Name = "MnuBodega";
			// 
			// mnuBodegaTRM
			// 
			this.mnuBodegaTRM.Caption = "Turnos de Recepción de Mercaderia";
			this.mnuBodegaTRM.Id = 856;
			this.mnuBodegaTRM.Name = "mnuBodegaTRM";
			this.mnuBodegaTRM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaTRM_ItemClick);
			// 
			// MnuFactSerie
			// 
			this.MnuFactSerie.Caption = "Entrega y recepción de mercadería";
			this.MnuFactSerie.Id = 455;
			this.MnuFactSerie.Name = "MnuFactSerie";
			this.MnuFactSerie.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactSerie_ItemClick);
			// 
			// MnuFacInv
			// 
			this.MnuFacInv.Caption = "Inventario";
			this.MnuFacInv.Id = 484;
			this.MnuFacInv.ItemShortcut = new DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F4);
			this.MnuFacInv.Name = "MnuFacInv";
			this.MnuFacInv.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFacInv_ItemClick);
			// 
			// mnuBodegaKardex
			// 
			this.mnuBodegaKardex.Caption = "kárdex";
			this.mnuBodegaKardex.Id = 540;
			this.mnuBodegaKardex.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.K));
			this.mnuBodegaKardex.Name = "mnuBodegaKardex";
			this.mnuBodegaKardex.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaKardex_ItemClick);
			// 
			// MnuFactLoteSpif
			// 
			this.MnuFactLoteSpif.Caption = "SPIF";
			this.MnuFactLoteSpif.Id = 473;
			this.MnuFactLoteSpif.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S));
			this.MnuFactLoteSpif.Name = "MnuFactLoteSpif";
			this.MnuFactLoteSpif.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactLoteSpif_ItemClick);
			// 
			// MnuFacSeriales
			// 
			this.MnuFacSeriales.Caption = "Movimiento de Seriales";
			this.MnuFacSeriales.Id = 507;
			this.MnuFacSeriales.Name = "MnuFacSeriales";
			this.MnuFacSeriales.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFacSeriales_ItemClick);
			// 
			// mnuBodegaIngSeriales
			// 
			this.mnuBodegaIngSeriales.Caption = "Ingreso de Seriales";
			this.mnuBodegaIngSeriales.Id = 660;
			this.mnuBodegaIngSeriales.Name = "mnuBodegaIngSeriales";
			this.mnuBodegaIngSeriales.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaIngSeriales_ItemClick);
			// 
			// mnuBodegaIPR
			// 
			this.mnuBodegaIPR.Caption = "Ingreso de Productos de Remate";
			this.mnuBodegaIPR.Id = 734;
			this.mnuBodegaIPR.Name = "mnuBodegaIPR";
			this.mnuBodegaIPR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaIPR_ItemClick);
			// 
			// mnuFactTransf
			// 
			this.mnuFactTransf.Caption = "Transferencias";
			this.mnuFactTransf.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.mnuFactTransf.Id = 211;
			this.mnuFactTransf.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.T));
			this.mnuFactTransf.Name = "mnuFactTransf";
			this.mnuFactTransf.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFactTransf_ItemClick);
			// 
			// MnuFactLoteTransMult
			// 
			this.MnuFactLoteTransMult.Caption = "Transferencia Multiple";
			this.MnuFactLoteTransMult.Id = 472;
			this.MnuFactLoteTransMult.Name = "MnuFactLoteTransMult";
			this.MnuFactLoteTransMult.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactLoteTransMult_ItemClick);
			// 
			// mnuComprasTEL
			// 
			this.mnuComprasTEL.Caption = "Transferencias Entre Locales";
			this.mnuComprasTEL.Id = 611;
			this.mnuComprasTEL.Name = "mnuComprasTEL";
			this.mnuComprasTEL.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasTEL_ItemClick);
			// 
			// MnuFactEstTransf
			// 
			this.MnuFactEstTransf.Caption = "Estado Transferencias";
			this.MnuFactEstTransf.Id = 510;
			this.MnuFactEstTransf.Name = "MnuFactEstTransf";
			this.MnuFactEstTransf.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactEstTransf_ItemClick);
			// 
			// mnuFactGuiaRemision
			// 
			this.mnuFactGuiaRemision.Caption = "Guia de Remisión";
			this.mnuFactGuiaRemision.Id = 511;
			this.mnuFactGuiaRemision.Name = "mnuFactGuiaRemision";
			this.mnuFactGuiaRemision.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuFactGuiaRemision.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFactGuiaRemision_ItemClick);
			// 
			// mnuBodegaGRC
			// 
			this.mnuBodegaGRC.Caption = "Guia de Remisión de Carpa";
			this.mnuBodegaGRC.Id = 664;
			this.mnuBodegaGRC.Name = "mnuBodegaGRC";
			this.mnuBodegaGRC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaGRC_ItemClick);
			// 
			// MnuFactEnvTransferencia
			// 
			this.MnuFactEnvTransferencia.Caption = "Egreso de Bodega";
			this.MnuFactEnvTransferencia.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactEnvTransferencia.Id = 111;
			this.MnuFactEnvTransferencia.ImageIndex = 33;
			this.MnuFactEnvTransferencia.Name = "MnuFactEnvTransferencia";
			this.MnuFactEnvTransferencia.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactEnvTransferencia_ItemClick);
			// 
			// MnuFactRecTransferencia
			// 
			this.MnuFactRecTransferencia.Caption = "Ingreso de Bodega";
			this.MnuFactRecTransferencia.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactRecTransferencia.Id = 112;
			this.MnuFactRecTransferencia.ImageIndex = 19;
			this.MnuFactRecTransferencia.Name = "MnuFactRecTransferencia";
			this.MnuFactRecTransferencia.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactRecTransferencia_ItemClick);
			// 
			// mnuBodegaNI
			// 
			this.mnuBodegaNI.Caption = "Nota De Ingreso";
			this.mnuBodegaNI.Id = 975;
			this.mnuBodegaNI.Name = "mnuBodegaNI";
			this.mnuBodegaNI.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaNI_ItemClick);
			// 
			// mnuBodegaCDI
			// 
			this.mnuBodegaCDI.Caption = "Conteo Diario de Inventario";
			this.mnuBodegaCDI.Id = 799;
			this.mnuBodegaCDI.Name = "mnuBodegaCDI";
			this.mnuBodegaCDI.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaCDI_ItemClick);
			// 
			// mnuBodegaRM
			// 
			this.mnuBodegaRM.Caption = "Requerimiento de Mercaderia";
			this.mnuBodegaRM.Id = 788;
			this.mnuBodegaRM.Name = "mnuBodegaRM";
			this.mnuBodegaRM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaRM_ItemClick);
			// 
			// mnuBodegaCI
			// 
			this.mnuBodegaCI.Caption = "Conteo de Inventario";
			this.mnuBodegaCI.Id = 569;
			this.mnuBodegaCI.Name = "mnuBodegaCI";
			this.mnuBodegaCI.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaCI_ItemClick);
			// 
			// mnuBodegaDCI
			// 
			this.mnuBodegaDCI.Caption = "Diferencias Conteo Inventario";
			this.mnuBodegaDCI.Id = 578;
			this.mnuBodegaDCI.Name = "mnuBodegaDCI";
			this.mnuBodegaDCI.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaDCI_ItemClick);
			// 
			// mnuBodegaCierreInventario
			// 
			this.mnuBodegaCierreInventario.Caption = "Cierre Conteo Inventario";
			this.mnuBodegaCierreInventario.Id = 591;
			this.mnuBodegaCierreInventario.Name = "mnuBodegaCierreInventario";
			this.mnuBodegaCierreInventario.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaCierreInventario_ItemClick);
			// 
			// mnuReportesBodega
			// 
			this.mnuReportesBodega.Caption = "Reportes";
			this.mnuReportesBodega.Id = 722;
			this.mnuReportesBodega.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																	 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaIF),
																																																	 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaInvVal),
																																																	 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRepCmpTransf),
																																																	 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaRepKS),
																																																	 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaReportesECDI),
																																																	 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaRepMD),
																																																	 new DevExpress.XtraBars.LinkPersistInfo(this.mnuBodegaReportesKardexCosto)});
			this.mnuReportesBodega.Name = "mnuReportesBodega";
			// 
			// mnuBodegaIF
			// 
			this.mnuBodegaIF.Caption = "Inventario Por Fecha";
			this.mnuBodegaIF.Id = 604;
			this.mnuBodegaIF.Name = "mnuBodegaIF";
			this.mnuBodegaIF.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaIF_ItemClick);
			// 
			// mnuBodegaInvVal
			// 
			this.mnuBodegaInvVal.Caption = "Inventario Valorado";
			this.mnuBodegaInvVal.Id = 614;
			this.mnuBodegaInvVal.Name = "mnuBodegaInvVal";
			this.mnuBodegaInvVal.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaInvVal_ItemClick);
			// 
			// mnuRepCmpTransf
			// 
			this.mnuRepCmpTransf.Caption = "Reporte de Transferencias";
			this.mnuRepCmpTransf.Id = 589;
			this.mnuRepCmpTransf.Name = "mnuRepCmpTransf";
			this.mnuRepCmpTransf.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRepCmpTransf_ItemClick);
			// 
			// mnuBodegaRepKS
			// 
			this.mnuBodegaRepKS.Caption = "Kardex de Seriales";
			this.mnuBodegaRepKS.Id = 790;
			this.mnuBodegaRepKS.Name = "mnuBodegaRepKS";
			this.mnuBodegaRepKS.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaRepKS_ItemClick);
			// 
			// mnuBodegaReportesECDI
			// 
			this.mnuBodegaReportesECDI.Caption = "Estado de Conteos de Inventarios";
			this.mnuBodegaReportesECDI.Id = 817;
			this.mnuBodegaReportesECDI.Name = "mnuBodegaReportesECDI";
			this.mnuBodegaReportesECDI.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaReportesECDI_ItemClick);
			// 
			// mnuBodegaRepMD
			// 
			this.mnuBodegaRepMD.Caption = "Movimientos Descuadrados";
			this.mnuBodegaRepMD.Id = 853;
			this.mnuBodegaRepMD.Name = "mnuBodegaRepMD";
			this.mnuBodegaRepMD.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaRepMD_ItemClick);
			// 
			// mnuBodegaReportesKardexCosto
			// 
			this.mnuBodegaReportesKardexCosto.Caption = "Kardex Con Costo";
			this.mnuBodegaReportesKardexCosto.Id = 895;
			this.mnuBodegaReportesKardexCosto.Name = "mnuBodegaReportesKardexCosto";
			this.mnuBodegaReportesKardexCosto.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaReportesKardexCosto_ItemClick);
			// 
			// mnuFactBodegas
			// 
			this.mnuFactBodegas.Caption = "Bodegas";
			this.mnuFactBodegas.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuFactBodegas.Id = 306;
			this.mnuFactBodegas.Name = "mnuFactBodegas";
			this.mnuFactBodegas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFactBodegas_ItemClick);
			// 
			// mnuBodegaSectorizacion
			// 
			this.mnuBodegaSectorizacion.Caption = "Sectorizacion";
			this.mnuBodegaSectorizacion.Id = 610;
			this.mnuBodegaSectorizacion.Name = "mnuBodegaSectorizacion";
			this.mnuBodegaSectorizacion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuBodegaSectorizacion_ItemClick);
			// 
			// MnuHerr_
			// 
			this.MnuHerr_.Caption = "&Herramientas";
			this.MnuHerr_.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuHerr_.Id = 11;
			this.MnuHerr_.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																													new DevExpress.XtraBars.LinkPersistInfo(this.MnuHerr_Camb_Contr),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.MnuHerr_Actu_Vers, true),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuHerr_Sal_Ini, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuHerr_Camb_Ctes, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuHerr_Copi_Arch, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuHerr_Auditoría, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuHerr_Act_Saldo, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.MnuEdicSeguridad),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerrGenSQL, "", false, true, false, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerrTester, "", false, true, false, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerrReporte, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuFactRepComp, "", false, true, false, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerrCredito, "", false, true, false, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerrBusqueda, "", false, true, false, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerrContratos, "", false, true, false, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerrPresupC, "", false, true, false, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerramientasNumeracion, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerramientasCRIVA, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerrALF, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerramientasPTP, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerramientasAP, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerramientasPTA, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerramientasAPL, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuHerramientasASDP, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuVentasEOC, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuVentasGCR, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrVerifDepositos),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrSegClientes),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrDepositos),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuSolNC)});
			this.MnuHerr_.Name = "MnuHerr_";
			// 
			// MnuHerr_Camb_Contr
			// 
			this.MnuHerr_Camb_Contr.Caption = "Cambiar contraseña";
			this.MnuHerr_Camb_Contr.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuHerr_Camb_Contr.Id = 65;
			this.MnuHerr_Camb_Contr.Name = "MnuHerr_Camb_Contr";
			this.MnuHerr_Camb_Contr.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuHerr_Camb_Contr_ItemClick);
			// 
			// MnuHerr_Actu_Vers
			// 
			this.MnuHerr_Actu_Vers.Caption = "Actualizar Versión";
			this.MnuHerr_Actu_Vers.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuHerr_Actu_Vers.Id = 67;
			this.MnuHerr_Actu_Vers.ImageIndex = 43;
			this.MnuHerr_Actu_Vers.Name = "MnuHerr_Actu_Vers";
			this.MnuHerr_Actu_Vers.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuHerr_Actu_Vers_ItemClick);
			// 
			// MnuHerr_Sal_Ini
			// 
			this.MnuHerr_Sal_Ini.Caption = "Saldos Iniciales";
			this.MnuHerr_Sal_Ini.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuHerr_Sal_Ini.Id = 70;
			this.MnuHerr_Sal_Ini.Name = "MnuHerr_Sal_Ini";
			this.MnuHerr_Sal_Ini.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuHerr_Sal_Ini_ItemClick);
			// 
			// MnuHerr_Camb_Ctes
			// 
			this.MnuHerr_Camb_Ctes.Caption = "Cambio de Constantes";
			this.MnuHerr_Camb_Ctes.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuHerr_Camb_Ctes.Id = 74;
			this.MnuHerr_Camb_Ctes.Name = "MnuHerr_Camb_Ctes";
			this.MnuHerr_Camb_Ctes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuHerr_Camb_Ctes_ItemClick);
			// 
			// MnuHerr_Copi_Arch
			// 
			this.MnuHerr_Copi_Arch.Caption = "Copiar Archivos";
			this.MnuHerr_Copi_Arch.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuHerr_Copi_Arch.Id = 75;
			this.MnuHerr_Copi_Arch.Name = "MnuHerr_Copi_Arch";
			this.MnuHerr_Copi_Arch.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuHerr_Copi_Arch_ItemClick);
			// 
			// MnuHerr_Auditoría
			// 
			this.MnuHerr_Auditoría.Caption = "Auditoría";
			this.MnuHerr_Auditoría.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuHerr_Auditoría.Id = 77;
			this.MnuHerr_Auditoría.Name = "MnuHerr_Auditoría";
			this.MnuHerr_Auditoría.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuHerr_Auditoría_ItemClick);
			// 
			// MnuHerr_Act_Saldo
			// 
			this.MnuHerr_Act_Saldo.Caption = "Actualiza Saldos";
			this.MnuHerr_Act_Saldo.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuHerr_Act_Saldo.Id = 78;
			this.MnuHerr_Act_Saldo.Name = "MnuHerr_Act_Saldo";
			this.MnuHerr_Act_Saldo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuHerr_Resumen_ItemClick);
			// 
			// MnuEdicSeguridad
			// 
			this.MnuEdicSeguridad.Caption = "Seguridad";
			this.MnuEdicSeguridad.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuEdicSeguridad.Id = 53;
			this.MnuEdicSeguridad.ImageIndex = 13;
			this.MnuEdicSeguridad.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.MnuEdicSegMenu),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.MnuEdicSegUsuario),
																																																	new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuHerrSegGenGrupos, false),
																																																	new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.MnuEdicSegConfigUsuarios, false),
																																																	new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuHerrSegFiltro, false),
																																																	new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuHerrSegFiltroSuc, false),
																																																	new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mbuHerrSegTransf, false),
																																																	new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuHerrSegApReq, false),
																																																	new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuHerrSegReqLis, false),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrSegCD),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem68)});
			this.MnuEdicSeguridad.Name = "MnuEdicSeguridad";
			// 
			// MnuEdicSegMenu
			// 
			this.MnuEdicSegMenu.Caption = "Menú";
			this.MnuEdicSegMenu.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuEdicSegMenu.Id = 54;
			this.MnuEdicSegMenu.Name = "MnuEdicSegMenu";
			this.MnuEdicSegMenu.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.Mnu_EdicSeguridad_ItemClick);
			// 
			// MnuEdicSegUsuario
			// 
			this.MnuEdicSegUsuario.Caption = "Usuarios";
			this.MnuEdicSegUsuario.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuEdicSegUsuario.Id = 55;
			this.MnuEdicSegUsuario.Name = "MnuEdicSegUsuario";
			this.MnuEdicSegUsuario.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.bbiUsua_ItemClick);
			// 
			// mnuHerrSegGenGrupos
			// 
			this.mnuHerrSegGenGrupos.Caption = "Especificaciones de Grupos";
			this.mnuHerrSegGenGrupos.Id = 440;
			this.mnuHerrSegGenGrupos.Name = "mnuHerrSegGenGrupos";
			this.mnuHerrSegGenGrupos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuHerrSegGenGrupos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrSegGenGrupos_ItemClick);
			// 
			// MnuEdicSegConfigUsuarios
			// 
			this.MnuEdicSegConfigUsuarios.Caption = "Configurar Cuentas";
			this.MnuEdicSegConfigUsuarios.Id = 442;
			this.MnuEdicSegConfigUsuarios.Name = "MnuEdicSegConfigUsuarios";
			this.MnuEdicSegConfigUsuarios.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.MnuEdicSegConfigUsuarios.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuEdicSegConfigUsuarios_ItemClick);
			// 
			// mnuHerrSegFiltro
			// 
			this.mnuHerrSegFiltro.Caption = "Filtro Artículos";
			this.mnuHerrSegFiltro.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuHerrSegFiltro.Id = 305;
			this.mnuHerrSegFiltro.Name = "mnuHerrSegFiltro";
			this.mnuHerrSegFiltro.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuHerrSegFiltro.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrSegFiltro_ItemClick);
			// 
			// mnuHerrSegFiltroSuc
			// 
			this.mnuHerrSegFiltroSuc.Caption = "Filtro Sucursal";
			this.mnuHerrSegFiltroSuc.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuHerrSegFiltroSuc.Id = 310;
			this.mnuHerrSegFiltroSuc.Name = "mnuHerrSegFiltroSuc";
			this.mnuHerrSegFiltroSuc.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuHerrSegFiltroSuc.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrSegFiltroSuc_ItemClick);
			// 
			// mbuHerrSegTransf
			// 
			this.mbuHerrSegTransf.Caption = "Cupo Transformación";
			this.mbuHerrSegTransf.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mbuHerrSegTransf.Id = 311;
			this.mbuHerrSegTransf.Name = "mbuHerrSegTransf";
			this.mbuHerrSegTransf.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mbuHerrSegTransf.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mbuHerrSegTransf_ItemClick);
			// 
			// mnuHerrSegApReq
			// 
			this.mnuHerrSegApReq.Caption = "Requisición Orden C. ";
			this.mnuHerrSegApReq.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuHerrSegApReq.Id = 384;
			this.mnuHerrSegApReq.Name = "mnuHerrSegApReq";
			this.mnuHerrSegApReq.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuHerrSegApReq.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrSegApReq_ItemClick);
			// 
			// mnuHerrSegReqLis
			// 
			this.mnuHerrSegReqLis.Caption = "Requisición No Aprob.";
			this.mnuHerrSegReqLis.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuHerrSegReqLis.Id = 397;
			this.mnuHerrSegReqLis.Name = "mnuHerrSegReqLis";
			this.mnuHerrSegReqLis.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuHerrSegReqLis.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrSegReqLis_ItemClick);
			// 
			// mnuHerrSegCD
			// 
			this.mnuHerrSegCD.Caption = "Clave Descuentos";
			this.mnuHerrSegCD.Id = 586;
			this.mnuHerrSegCD.Name = "mnuHerrSegCD";
			this.mnuHerrSegCD.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrSegCD_ItemClick);
			// 
			// barButtonItem68
			// 
			this.barButtonItem68.Caption = "utilitarios";
			this.barButtonItem68.Id = 769;
			this.barButtonItem68.Name = "barButtonItem68";
			this.barButtonItem68.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem68_ItemClick);
			// 
			// mnuHerrGenSQL
			// 
			this.mnuHerrGenSQL.Caption = "Generar Sql";
			this.mnuHerrGenSQL.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuHerrGenSQL.Id = 237;
			this.mnuHerrGenSQL.Name = "mnuHerrGenSQL";
			this.mnuHerrGenSQL.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuHerrGenSQL.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrGenSQL_ItemClick);
			// 
			// mnuHerrTester
			// 
			this.mnuHerrTester.Caption = "Tester";
			this.mnuHerrTester.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuHerrTester.Id = 281;
			this.mnuHerrTester.Name = "mnuHerrTester";
			this.mnuHerrTester.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuHerrTester.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btHerrTester_ItemClick);
			// 
			// mnuHerrReporte
			// 
			this.mnuHerrReporte.Caption = "Reportes";
			this.mnuHerrReporte.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuHerrReporte.Id = 243;
			this.mnuHerrReporte.Name = "mnuHerrReporte";
			this.mnuHerrReporte.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuHerrReporte.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrReporte_ItemClick);
			// 
			// mnuFactRepComp
			// 
			this.mnuFactRepComp.Caption = "Inventario Avanzado";
			this.mnuFactRepComp.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuFactRepComp.Id = 391;
			this.mnuFactRepComp.Name = "mnuFactRepComp";
			this.mnuFactRepComp.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuFactRepComp.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFactRepComp_ItemClick);
			// 
			// mnuHerrCredito
			// 
			this.mnuHerrCredito.Caption = "Análisis de Crédito";
			this.mnuHerrCredito.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuHerrCredito.Id = 392;
			this.mnuHerrCredito.Name = "mnuHerrCredito";
			this.mnuHerrCredito.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuHerrCredito.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrCredito_ItemClick);
			// 
			// mnuHerrBusqueda
			// 
			this.mnuHerrBusqueda.Caption = "Búsqueda";
			this.mnuHerrBusqueda.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuHerrBusqueda.Id = 395;
			this.mnuHerrBusqueda.Name = "mnuHerrBusqueda";
			this.mnuHerrBusqueda.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuHerrBusqueda.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrBusqueda_ItemClick);
			// 
			// mnuHerrContratos
			// 
			this.mnuHerrContratos.Caption = "Contratos";
			this.mnuHerrContratos.Enabled = false;
			this.mnuHerrContratos.Id = 407;
			this.mnuHerrContratos.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrContratosIngreso),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrContratosGeneraFact),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrContratosAmortiza),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrContratosAdemdums),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrContratosAlarma)});
			this.mnuHerrContratos.Name = "mnuHerrContratos";
			this.mnuHerrContratos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuHerrContratosIngreso
			// 
			this.mnuHerrContratosIngreso.Caption = "Ingreso Contratos";
			this.mnuHerrContratosIngreso.Id = 408;
			this.mnuHerrContratosIngreso.Name = "mnuHerrContratosIngreso";
			this.mnuHerrContratosIngreso.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrContratosIngreso_ItemClick);
			// 
			// mnuHerrContratosGeneraFact
			// 
			this.mnuHerrContratosGeneraFact.Caption = "Generar Facturas";
			this.mnuHerrContratosGeneraFact.Id = 409;
			this.mnuHerrContratosGeneraFact.Name = "mnuHerrContratosGeneraFact";
			this.mnuHerrContratosGeneraFact.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrContratosGeneraFact_ItemClick);
			// 
			// mnuHerrContratosAmortiza
			// 
			this.mnuHerrContratosAmortiza.Caption = "Tabla Amortizacion";
			this.mnuHerrContratosAmortiza.Id = 453;
			this.mnuHerrContratosAmortiza.Name = "mnuHerrContratosAmortiza";
			this.mnuHerrContratosAmortiza.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrContratosAmortiza_ItemClick);
			// 
			// mnuHerrContratosAdemdums
			// 
			this.mnuHerrContratosAdemdums.Caption = "Ademdums";
			this.mnuHerrContratosAdemdums.Id = 414;
			this.mnuHerrContratosAdemdums.Name = "mnuHerrContratosAdemdums";
			this.mnuHerrContratosAdemdums.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuHerrContratosAdemdums.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrContratosAdemdums_ItemClick);
			// 
			// mnuHerrContratosAlarma
			// 
			this.mnuHerrContratosAlarma.Caption = "Alarma";
			this.mnuHerrContratosAlarma.Id = 413;
			this.mnuHerrContratosAlarma.Name = "mnuHerrContratosAlarma";
			this.mnuHerrContratosAlarma.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrContratosAlarma_ItemClick);
			// 
			// mnuHerrPresupC
			// 
			this.mnuHerrPresupC.Caption = "Presupuesto Compras";
			this.mnuHerrPresupC.Id = 429;
			this.mnuHerrPresupC.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrPresupCIngreso),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrPresupCAvanceObra),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrPresupCSolicitudPago),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrPresupCAnalisis, true),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrPresupCReportes),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuHerrPresupCImportar, true)});
			this.mnuHerrPresupC.Name = "mnuHerrPresupC";
			this.mnuHerrPresupC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuHerrPresupCIngreso
			// 
			this.mnuHerrPresupCIngreso.Caption = "Ingreso Presupuesto";
			this.mnuHerrPresupCIngreso.Id = 430;
			this.mnuHerrPresupCIngreso.Name = "mnuHerrPresupCIngreso";
			this.mnuHerrPresupCIngreso.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrPresupCIngreso_ItemClick);
			// 
			// mnuHerrPresupCAvanceObra
			// 
			this.mnuHerrPresupCAvanceObra.Caption = "Avance de Obra";
			this.mnuHerrPresupCAvanceObra.Id = 434;
			this.mnuHerrPresupCAvanceObra.Name = "mnuHerrPresupCAvanceObra";
			this.mnuHerrPresupCAvanceObra.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrPresupCAvanceObra_ItemClick);
			// 
			// mnuHerrPresupCSolicitudPago
			// 
			this.mnuHerrPresupCSolicitudPago.Caption = "Solicitud de Pago";
			this.mnuHerrPresupCSolicitudPago.Id = 450;
			this.mnuHerrPresupCSolicitudPago.Name = "mnuHerrPresupCSolicitudPago";
			this.mnuHerrPresupCSolicitudPago.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrPresupCSolicitudPago_ItemClick);
			// 
			// mnuHerrPresupCAnalisis
			// 
			this.mnuHerrPresupCAnalisis.Caption = "Analisis Costos";
			this.mnuHerrPresupCAnalisis.Id = 431;
			this.mnuHerrPresupCAnalisis.Name = "mnuHerrPresupCAnalisis";
			this.mnuHerrPresupCAnalisis.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrPresupCAnalisis_ItemClick);
			// 
			// mnuHerrPresupCReportes
			// 
			this.mnuHerrPresupCReportes.Caption = "Reportes";
			this.mnuHerrPresupCReportes.Id = 451;
			this.mnuHerrPresupCReportes.Name = "mnuHerrPresupCReportes";
			this.mnuHerrPresupCReportes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrPresupCReportes_ItemClick);
			// 
			// mnuHerrPresupCImportar
			// 
			this.mnuHerrPresupCImportar.Caption = "Importar Presupuesto";
			this.mnuHerrPresupCImportar.Id = 435;
			this.mnuHerrPresupCImportar.Name = "mnuHerrPresupCImportar";
			this.mnuHerrPresupCImportar.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrPresupCImportar_ItemClick);
			// 
			// mnuHerramientasNumeracion
			// 
			this.mnuHerramientasNumeracion.Caption = "Numeración de los locales";
			this.mnuHerramientasNumeracion.Id = 552;
			this.mnuHerramientasNumeracion.Name = "mnuHerramientasNumeracion";
			this.mnuHerramientasNumeracion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerramientasNumeracion_ItemClick);
			// 
			// mnuHerramientasCRIVA
			// 
			this.mnuHerramientasCRIVA.Caption = "Retenciones IVA";
			this.mnuHerramientasCRIVA.Id = 601;
			this.mnuHerramientasCRIVA.Name = "mnuHerramientasCRIVA";
			this.mnuHerramientasCRIVA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerramientasCRIVA_ItemClick);
			// 
			// mnuHerrALF
			// 
			this.mnuHerrALF.Caption = "Anula Facturas en Lote";
			this.mnuHerrALF.Id = 659;
			this.mnuHerrALF.Name = "mnuHerrALF";
			this.mnuHerrALF.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrALF_ItemClick);
			// 
			// mnuHerramientasPTP
			// 
			this.mnuHerramientasPTP.Caption = "Planificacion Tareas Personal";
			this.mnuHerramientasPTP.Id = 773;
			this.mnuHerramientasPTP.Name = "mnuHerramientasPTP";
			this.mnuHerramientasPTP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerramientasPTP_ItemClick);
			// 
			// mnuHerramientasAP
			// 
			this.mnuHerramientasAP.Caption = "Asignacion Personal";
			this.mnuHerramientasAP.Id = 774;
			this.mnuHerramientasAP.Name = "mnuHerramientasAP";
			this.mnuHerramientasAP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerramientasAP_ItemClick);
			// 
			// mnuHerramientasPTA
			// 
			this.mnuHerramientasPTA.Caption = "Planificacion Tareas Almacenes";
			this.mnuHerramientasPTA.Id = 870;
			this.mnuHerramientasPTA.Name = "mnuHerramientasPTA";
			this.mnuHerramientasPTA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerramientasPTA_ItemClick);
			// 
			// mnuHerramientasAPL
			// 
			this.mnuHerramientasAPL.Caption = "Asignacion Personal Matriz";
			this.mnuHerramientasAPL.Id = 881;
			this.mnuHerramientasAPL.Name = "mnuHerramientasAPL";
			this.mnuHerramientasAPL.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerramientasAPL_ItemClick);
			// 
			// mnuHerramientasASDP
			// 
			this.mnuHerramientasASDP.Caption = "Servicios Por Departamentos";
			this.mnuHerramientasASDP.Id = 882;
			this.mnuHerramientasASDP.Name = "mnuHerramientasASDP";
			this.mnuHerramientasASDP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerramientasASDP_ItemClick);
			// 
			// mnuVentasEOC
			// 
			this.mnuVentasEOC.Caption = "Entrega de Obsequios Clientes";
			this.mnuVentasEOC.Id = 823;
			this.mnuVentasEOC.Name = "mnuVentasEOC";
			this.mnuVentasEOC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasEOC_ItemClick);
			// 
			// mnuVentasGCR
			// 
			this.mnuVentasGCR.Caption = "Gestiones de Clientes Referidos";
			this.mnuVentasGCR.Id = 826;
			this.mnuVentasGCR.Name = "mnuVentasGCR";
			this.mnuVentasGCR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasGCR_ItemClick);
			// 
			// mnuHerrVerifDepositos
			// 
			this.mnuHerrVerifDepositos.Caption = "Verificacion de Depositos";
			this.mnuHerrVerifDepositos.Id = 842;
			this.mnuHerrVerifDepositos.Name = "mnuHerrVerifDepositos";
			this.mnuHerrVerifDepositos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrVerifDepositos_ItemClick);
			// 
			// mnuHerrSegClientes
			// 
			this.mnuHerrSegClientes.Caption = "Seguimiento Clientes";
			this.mnuHerrSegClientes.Id = 843;
			this.mnuHerrSegClientes.Name = "mnuHerrSegClientes";
			this.mnuHerrSegClientes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrSegClientes_ItemClick);
			// 
			// mnuHerrDepositos
			// 
			this.mnuHerrDepositos.Caption = "Depositos ";
			this.mnuHerrDepositos.Id = 844;
			this.mnuHerrDepositos.Name = "mnuHerrDepositos";
			this.mnuHerrDepositos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrDepositos_ItemClick);
			// 
			// mnuSolNC
			// 
			this.mnuSolNC.Caption = "Solicitud De Notas De Crédito";
			this.mnuSolNC.Id = 992;
			this.mnuSolNC.Name = "mnuSolNC";
			this.mnuSolNC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuSolNC_ItemClick);
			// 
			// MnuCont
			// 
			this.MnuCont.Caption = "&Contabillidad";
			this.MnuCont.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuCont.Id = 82;
			this.MnuCont.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContPlanCtas),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContAsigCtas),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContAsientos, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFacArtServicios),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactProyectos),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContFlujo, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContTCC),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContConsumosTC),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContLCTC),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.barSubItem4, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContSRI),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactAutPagos, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFacPagosAut),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactActFijos, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContBloqAsientos, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContCierPeriodo),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContReportes, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuCont_Ind_Financieros, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContPresupuesto),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContLote),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContTransacciones, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContRepCC),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContAsigArt),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContabilidadProvisiones),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContRB),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContCaja)});
			this.MnuCont.Name = "MnuCont";
			// 
			// MnuContPlanCtas
			// 
			this.MnuContPlanCtas.Caption = "Plan de Cuentas";
			this.MnuContPlanCtas.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContPlanCtas.Id = 83;
			this.MnuContPlanCtas.ImageIndex = 14;
			this.MnuContPlanCtas.Name = "MnuContPlanCtas";
			this.MnuContPlanCtas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCont_Plan_Ctas_ItemClick);
			// 
			// MnuContAsigCtas
			// 
			this.MnuContAsigCtas.Caption = "Asignación de Cuentas";
			this.MnuContAsigCtas.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContAsigCtas.Id = 100;
			this.MnuContAsigCtas.Name = "MnuContAsigCtas";
			this.MnuContAsigCtas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCont_Asig_Ctas_ItemClick);
			// 
			// MnuContAsientos
			// 
			this.MnuContAsientos.Caption = "Asientos      ";
			this.MnuContAsientos.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContAsientos.Id = 84;
			this.MnuContAsientos.ImageIndex = 36;
			this.MnuContAsientos.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.A));
			this.MnuContAsientos.Name = "MnuContAsientos";
			this.MnuContAsientos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCont_Asientos_ItemClick);
			// 
			// MnuFacArtServicios
			// 
			this.MnuFacArtServicios.Caption = "Servicios";
			this.MnuFacArtServicios.Id = 500;
			this.MnuFacArtServicios.Name = "MnuFacArtServicios";
			this.MnuFacArtServicios.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFacArtServicios_ItemClick);
			// 
			// MnuFactProyectos
			// 
			this.MnuFactProyectos.Caption = "Centros de Costo";
			this.MnuFactProyectos.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactProyectos.Id = 130;
			this.MnuFactProyectos.Name = "MnuFactProyectos";
			this.MnuFactProyectos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactProyectos_ItemClick);
			// 
			// MnuContFlujo
			// 
			this.MnuContFlujo.Caption = "Flujo de Caja Chica";
			this.MnuContFlujo.Id = 402;
			this.MnuContFlujo.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F10));
			this.MnuContFlujo.Name = "MnuContFlujo";
			this.MnuContFlujo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContFlujo_ItemClick);
			// 
			// mnuContTCC
			// 
			this.mnuContTCC.Caption = "Tarjetas De Crédito Corporativas";
			this.mnuContTCC.Id = 581;
			this.mnuContTCC.Name = "mnuContTCC";
			this.mnuContTCC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContTCC_ItemClick);
			// 
			// mnuContConsumosTC
			// 
			this.mnuContConsumosTC.Caption = "Consumos con Tarjetas de Crédito";
			this.mnuContConsumosTC.Id = 577;
			this.mnuContConsumosTC.Name = "mnuContConsumosTC";
			this.mnuContConsumosTC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContConsumosTC_ItemClick);
			// 
			// mnuContLCTC
			// 
			this.mnuContLCTC.Caption = "Lista De Consumos con Tarjeta de Crédito";
			this.mnuContLCTC.Id = 579;
			this.mnuContLCTC.Name = "mnuContLCTC";
			this.mnuContLCTC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuContLCTC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContLCTC_ItemClick);
			// 
			// barSubItem4
			// 
			this.barSubItem4.Caption = "Bancos";
			this.barSubItem4.Id = 575;
			this.barSubItem4.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContBcos),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContCtas),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContMovBanc),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContEstadoCheques),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContTEB),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContPrestamos),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContEstadoCta, true),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContBancosCBA),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContBancosCP),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContBancosDPL, true),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContBancosCDPL)});
			this.barSubItem4.Name = "barSubItem4";
			// 
			// MnuContBcos
			// 
			this.MnuContBcos.Caption = "Bancos";
			this.MnuContBcos.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContBcos.Id = 96;
			this.MnuContBcos.Name = "MnuContBcos";
			this.MnuContBcos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCont_Bcos_ItemClick);
			// 
			// MnuContCtas
			// 
			this.MnuContCtas.Caption = "Cuentas Bancarias";
			this.MnuContCtas.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContCtas.Id = 95;
			this.MnuContCtas.Name = "MnuContCtas";
			this.MnuContCtas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCont_Ctas_ItemClick);
			// 
			// mnuContMovBanc
			// 
			this.mnuContMovBanc.Caption = "Movimientos Bancarios";
			this.mnuContMovBanc.Id = 565;
			this.mnuContMovBanc.Name = "mnuContMovBanc";
			this.mnuContMovBanc.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContMovBanc_ItemClick);
			// 
			// mnuContEstadoCheques
			// 
			this.mnuContEstadoCheques.Caption = "Estado Cheques";
			this.mnuContEstadoCheques.Id = 566;
			this.mnuContEstadoCheques.Name = "mnuContEstadoCheques";
			this.mnuContEstadoCheques.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContEstadoCheques_ItemClick);
			// 
			// mnuContTEB
			// 
			this.mnuContTEB.Caption = "Transferencias Entre Bancos";
			this.mnuContTEB.Id = 572;
			this.mnuContTEB.Name = "mnuContTEB";
			this.mnuContTEB.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContTEB_ItemClick);
			// 
			// mnuContPrestamos
			// 
			this.mnuContPrestamos.Caption = "Registro de Préstamos";
			this.mnuContPrestamos.Id = 571;
			this.mnuContPrestamos.Name = "mnuContPrestamos";
			this.mnuContPrestamos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContPrestamos_ItemClick);
			// 
			// MnuContEstadoCta
			// 
			this.MnuContEstadoCta.Caption = "Conciliación Bancaria";
			this.MnuContEstadoCta.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContEstadoCta.Id = 93;
			this.MnuContEstadoCta.ImageIndex = 37;
			this.MnuContEstadoCta.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.E));
			this.MnuContEstadoCta.Name = "MnuContEstadoCta";
			this.MnuContEstadoCta.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContEstadoCta_ItemClick);
			// 
			// mnuContBancosCBA
			// 
			this.mnuContBancosCBA.Caption = "Conciliación Bancaria Automatica";
			this.mnuContBancosCBA.Id = 613;
			this.mnuContBancosCBA.Name = "mnuContBancosCBA";
			this.mnuContBancosCBA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContBancosCBA_ItemClick);
			// 
			// mnuContBancosCP
			// 
			this.mnuContBancosCP.Caption = "Cheques Posfechados";
			this.mnuContBancosCP.Id = 747;
			this.mnuContBancosCP.Name = "mnuContBancosCP";
			this.mnuContBancosCP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContBancosCP_ItemClick);
			// 
			// mnuContBancosDPL
			// 
			this.mnuContBancosDPL.Caption = "Depositos Por Liquidar";
			this.mnuContBancosDPL.Id = 751;
			this.mnuContBancosDPL.Name = "mnuContBancosDPL";
			this.mnuContBancosDPL.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContBancosDPL_ItemClick);
			// 
			// mnuContBancosCDPL
			// 
			this.mnuContBancosCDPL.Caption = "Liquidacion Depositos";
			this.mnuContBancosCDPL.Id = 752;
			this.mnuContBancosCDPL.Name = "mnuContBancosCDPL";
			this.mnuContBancosCDPL.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContBancosCDPL_ItemClick);
			// 
			// MnuContSRI
			// 
			this.MnuContSRI.Caption = "S.R.I.";
			this.MnuContSRI.Id = 532;
			this.MnuContSRI.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuContTransAnexoTr),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuContTransFormulario),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuContTransConciliacion)});
			this.MnuContSRI.Name = "MnuContSRI";
			// 
			// mnuContTransAnexoTr
			// 
			this.mnuContTransAnexoTr.Caption = "Anexo Transaccional";
			this.mnuContTransAnexoTr.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.mnuContTransAnexoTr.Id = 266;
			this.mnuContTransAnexoTr.Name = "mnuContTransAnexoTr";
			this.mnuContTransAnexoTr.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContTransAnexoTr_ItemClick);
			// 
			// mnuContTransFormulario
			// 
			this.mnuContTransFormulario.Caption = "Formularios S.R.I.";
			this.mnuContTransFormulario.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.mnuContTransFormulario.Id = 271;
			this.mnuContTransFormulario.Name = "mnuContTransFormulario";
			this.mnuContTransFormulario.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContTransFormulario_ItemClick);
			// 
			// mnuContTransConciliacion
			// 
			this.mnuContTransConciliacion.Caption = "Conciliación Tributaria";
			this.mnuContTransConciliacion.Id = 433;
			this.mnuContTransConciliacion.Name = "mnuContTransConciliacion";
			this.mnuContTransConciliacion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btConciliacion_ItemClick);
			// 
			// MnuFactAutPagos
			// 
			this.MnuFactAutPagos.Caption = "Autorización de Pagos";
			this.MnuFactAutPagos.Id = 505;
			this.MnuFactAutPagos.Name = "MnuFactAutPagos";
			this.MnuFactAutPagos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactAutPagos_ItemClick);
			// 
			// MnuFacPagosAut
			// 
			this.MnuFacPagosAut.Caption = "Pagos Autorizados";
			this.MnuFacPagosAut.Id = 504;
			this.MnuFacPagosAut.Name = "MnuFacPagosAut";
			this.MnuFacPagosAut.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFacPagosAut_ItemClick);
			// 
			// MnuFactActFijos
			// 
			this.MnuFactActFijos.Caption = "Activos Fijos";
			this.MnuFactActFijos.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactActFijos.Id = 129;
			this.MnuFactActFijos.Name = "MnuFactActFijos";
			this.MnuFactActFijos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactActFijos_ItemClick);
			// 
			// MnuContBloqAsientos
			// 
			this.MnuContBloqAsientos.Caption = "Bloquear Asientos";
			this.MnuContBloqAsientos.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContBloqAsientos.Id = 98;
			this.MnuContBloqAsientos.Name = "MnuContBloqAsientos";
			this.MnuContBloqAsientos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.MnuContBloqAsientos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCont_Bloq_Asi_ItemClick);
			// 
			// MnuContCierPeriodo
			// 
			this.MnuContCierPeriodo.Caption = "Cierre de Período";
			this.MnuContCierPeriodo.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuContCierPeriodo.Id = 99;
			this.MnuContCierPeriodo.Name = "MnuContCierPeriodo";
			this.MnuContCierPeriodo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCont_Cier_Per_ItemClick);
			// 
			// mnuContReportes
			// 
			this.mnuContReportes.Caption = "Reportes";
			this.mnuContReportes.Id = 576;
			this.mnuContReportes.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContRepCont),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContabilidadAsDesc),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContTransR),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContRepCajasChicas),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContRepGastos, true),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContReportesGA),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContRepFacElec, true),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContRepEstadoCB),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContRepPresupAlm)});
			this.mnuContReportes.Name = "mnuContReportes";
			// 
			// MnuContRepCont
			// 
			this.MnuContRepCont.Caption = "Contabilidad      ";
			this.MnuContRepCont.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContRepCont.Id = 85;
			this.MnuContRepCont.ImageIndex = 16;
			this.MnuContRepCont.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.B));
			this.MnuContRepCont.Name = "MnuContRepCont";
			this.MnuContRepCont.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContRepCont_ItemClick);
			// 
			// mnuContabilidadAsDesc
			// 
			this.mnuContabilidadAsDesc.Caption = "Asientos Descuadrados";
			this.mnuContabilidadAsDesc.Id = 560;
			this.mnuContabilidadAsDesc.Name = "mnuContabilidadAsDesc";
			this.mnuContabilidadAsDesc.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContabilidadAsDesc_ItemClick);
			// 
			// MnuContTransR
			// 
			this.MnuContTransR.Caption = "Retenciones";
			this.MnuContTransR.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContTransR.Id = 236;
			this.MnuContTransR.Name = "MnuContTransR";
			this.MnuContTransR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContTransR_ItemClick);
			// 
			// mnuContRepCajasChicas
			// 
			this.mnuContRepCajasChicas.Caption = "Cajas Chicas";
			this.mnuContRepCajasChicas.Id = 740;
			this.mnuContRepCajasChicas.Name = "mnuContRepCajasChicas";
			this.mnuContRepCajasChicas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContRepCajasChicas_ItemClick);
			// 
			// mnuContRepGastos
			// 
			this.mnuContRepGastos.Caption = "Gastos";
			this.mnuContRepGastos.Id = 619;
			this.mnuContRepGastos.Name = "mnuContRepGastos";
			this.mnuContRepGastos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContRepGastos_ItemClick);
			// 
			// mnuContReportesGA
			// 
			this.mnuContReportesGA.Caption = "Gastos Anuales";
			this.mnuContReportesGA.Id = 745;
			this.mnuContReportesGA.Name = "mnuContReportesGA";
			this.mnuContReportesGA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContReportesGA_ItemClick);
			// 
			// mnuContRepFacElec
			// 
			this.mnuContRepFacElec.Caption = "Facturación Electronica";
			this.mnuContRepFacElec.Id = 622;
			this.mnuContRepFacElec.Name = "mnuContRepFacElec";
			this.mnuContRepFacElec.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuContRepFacElec.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContRepFacElec_ItemClick);
			// 
			// mnuContRepEstadoCB
			// 
			this.mnuContRepEstadoCB.Caption = "Estado Conciliación Bancaria ";
			this.mnuContRepEstadoCB.Id = 874;
			this.mnuContRepEstadoCB.Name = "mnuContRepEstadoCB";
			this.mnuContRepEstadoCB.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContRepEstadoCB_ItemClick);
			// 
			// mnuContRepPresupAlm
			// 
			this.mnuContRepPresupAlm.Caption = "Presupuesto por Almacen";
			this.mnuContRepPresupAlm.Id = 880;
			this.mnuContRepPresupAlm.Name = "mnuContRepPresupAlm";
			this.mnuContRepPresupAlm.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContRepPresupAlm_ItemClick);
			// 
			// MnuCont_Ind_Financieros
			// 
			this.MnuCont_Ind_Financieros.Caption = "Índices Financieros";
			this.MnuCont_Ind_Financieros.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuCont_Ind_Financieros.Id = 102;
			this.MnuCont_Ind_Financieros.Name = "MnuCont_Ind_Financieros";
			this.MnuCont_Ind_Financieros.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCont_Ind_Financieros_ItemClick);
			// 
			// MnuContPresupuesto
			// 
			this.MnuContPresupuesto.Caption = "Presupuesto";
			this.MnuContPresupuesto.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContPresupuesto.Id = 94;
			this.MnuContPresupuesto.Name = "MnuContPresupuesto";
			this.MnuContPresupuesto.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContPresupuesto_ItemClick);
			// 
			// mnuContLote
			// 
			this.mnuContLote.Caption = "Lotes";
			this.mnuContLote.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.mnuContLote.Id = 245;
			this.mnuContLote.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContCompara),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContAsientoPeriodo),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContLoteDepura),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContLotAsiComprob),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContLotCuentaObligatoria)});
			this.mnuContLote.Name = "mnuContLote";
			// 
			// mnuContCompara
			// 
			this.mnuContCompara.Caption = "Comparación";
			this.mnuContCompara.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.mnuContCompara.Id = 244;
			this.mnuContCompara.Name = "mnuContCompara";
			this.mnuContCompara.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContCompara_ItemClick);
			// 
			// mnuContAsientoPeriodo
			// 
			this.mnuContAsientoPeriodo.Caption = "Asientos por Períodos";
			this.mnuContAsientoPeriodo.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.mnuContAsientoPeriodo.Id = 246;
			this.mnuContAsientoPeriodo.Name = "mnuContAsientoPeriodo";
			this.mnuContAsientoPeriodo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContAsientoPeriodo_ItemClick);
			// 
			// mnuContLoteDepura
			// 
			this.mnuContLoteDepura.Caption = "Depuración";
			this.mnuContLoteDepura.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.mnuContLoteDepura.Id = 250;
			this.mnuContLoteDepura.Name = "mnuContLoteDepura";
			this.mnuContLoteDepura.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContLoteDepura_ItemClick);
			// 
			// mnuContLotAsiComprob
			// 
			this.mnuContLotAsiComprob.Caption = "Asiento por Comprobante";
			this.mnuContLotAsiComprob.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.mnuContLotAsiComprob.Id = 399;
			this.mnuContLotAsiComprob.Name = "mnuContLotAsiComprob";
			this.mnuContLotAsiComprob.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContLotAsiComprob_ItemClick);
			// 
			// mnuContLotCuentaObligatoria
			// 
			this.mnuContLotCuentaObligatoria.Caption = "Cuentas Obligatorias";
			this.mnuContLotCuentaObligatoria.Id = 421;
			this.mnuContLotCuentaObligatoria.Name = "mnuContLotCuentaObligatoria";
			this.mnuContLotCuentaObligatoria.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btCuentaObligatoria_ItemClick);
			// 
			// MnuContTransacciones
			// 
			this.MnuContTransacciones.Caption = "Retenciones";
			this.MnuContTransacciones.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContTransacciones.Id = 221;
			this.MnuContTransacciones.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																			new DevExpress.XtraBars.LinkPersistInfo(this.MnuContTransE),
																																																			new DevExpress.XtraBars.LinkPersistInfo(this.MnuContTransI)});
			this.MnuContTransacciones.Name = "MnuContTransacciones";
			this.MnuContTransacciones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// MnuContTransE
			// 
			this.MnuContTransE.Caption = "Retenciones de Compra";
			this.MnuContTransE.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContTransE.Id = 222;
			this.MnuContTransE.Name = "MnuContTransE";
			this.MnuContTransE.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContTransE_ItemClick);
			// 
			// MnuContTransI
			// 
			this.MnuContTransI.Caption = "Retenciones de Venta";
			this.MnuContTransI.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContTransI.Id = 223;
			this.MnuContTransI.Name = "MnuContTransI";
			this.MnuContTransI.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContTransI_ItemClick);
			// 
			// mnuContRepCC
			// 
			this.mnuContRepCC.Caption = "Reporte de Flujo de Caja Chica";
			this.mnuContRepCC.Id = 563;
			this.mnuContRepCC.Name = "mnuContRepCC";
			this.mnuContRepCC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuContRepCC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContRepCC_ItemClick);
			// 
			// MnuContAsigArt
			// 
			this.MnuContAsigArt.Caption = "Asignación de Cuentas de Grupo de Artículos";
			this.MnuContAsigArt.Id = 506;
			this.MnuContAsigArt.Name = "MnuContAsigArt";
			this.MnuContAsigArt.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.MnuContAsigArt.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContAsigArt_ItemClick);
			// 
			// mnuContabilidadProvisiones
			// 
			this.mnuContabilidadProvisiones.Caption = "Provisiones";
			this.mnuContabilidadProvisiones.Id = 807;
			this.mnuContabilidadProvisiones.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																						new DevExpress.XtraBars.LinkPersistInfo(this.mnuContabilidadFlujoPMP),
																																																						new DevExpress.XtraBars.LinkPersistInfo(this.mnuContabilidadFPPP),
																																																						new DevExpress.XtraBars.LinkPersistInfo(this.mnuContFPPP)});
			this.mnuContabilidadProvisiones.Name = "mnuContabilidadProvisiones";
			// 
			// mnuContabilidadFlujoPMP
			// 
			this.mnuContabilidadFlujoPMP.Caption = "Configuración";
			this.mnuContabilidadFlujoPMP.Id = 808;
			this.mnuContabilidadFlujoPMP.Name = "mnuContabilidadFlujoPMP";
			this.mnuContabilidadFlujoPMP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContabilidadFlujoPMP_ItemClick);
			// 
			// mnuContabilidadFPPP
			// 
			this.mnuContabilidadFPPP.Caption = "Provisión";
			this.mnuContabilidadFPPP.Id = 845;
			this.mnuContabilidadFPPP.Name = "mnuContabilidadFPPP";
			this.mnuContabilidadFPPP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContabilidadFPPP_ItemClick);
			// 
			// mnuContFPPP
			// 
			this.mnuContFPPP.Caption = "Autorización de Pago A Proveedores";
			this.mnuContFPPP.Id = 809;
			this.mnuContFPPP.Name = "mnuContFPPP";
			this.mnuContFPPP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem53_ItemClick_1);
			// 
			// mnuContRB
			// 
			this.mnuContRB.Caption = "Retenciones Bancarias";
			this.mnuContRB.Id = 854;
			this.mnuContRB.Name = "mnuContRB";
			this.mnuContRB.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContRB_ItemClick);
			// 
			// mnuContCaja
			// 
			this.mnuContCaja.Caption = "Caja";
			this.mnuContCaja.Id = 875;
			this.mnuContCaja.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContEstadosCajas),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContIngresoDiarioCaja),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContIngresosDeCajas),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContEstadoGeneralDeCaja)});
			this.mnuContCaja.Name = "mnuContCaja";
			// 
			// mnuContEstadosCajas
			// 
			this.mnuContEstadosCajas.Caption = "Estado De Caja";
			this.mnuContEstadosCajas.Id = 876;
			this.mnuContEstadosCajas.Name = "mnuContEstadosCajas";
			this.mnuContEstadosCajas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContEstadosCajas_ItemClick);
			// 
			// mnuContIngresoDiarioCaja
			// 
			this.mnuContIngresoDiarioCaja.Caption = "Ingreso Diario De Caja";
			this.mnuContIngresoDiarioCaja.Id = 877;
			this.mnuContIngresoDiarioCaja.Name = "mnuContIngresoDiarioCaja";
			this.mnuContIngresoDiarioCaja.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContIngresoDiarioCaja_ItemClick);
			// 
			// mnuContIngresosDeCajas
			// 
			this.mnuContIngresosDeCajas.Caption = "Ingresos De Caja";
			this.mnuContIngresosDeCajas.Id = 878;
			this.mnuContIngresosDeCajas.Name = "mnuContIngresosDeCajas";
			this.mnuContIngresosDeCajas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContIngresosDeCajas_ItemClick);
			// 
			// mnuContEstadoGeneralDeCaja
			// 
			this.mnuContEstadoGeneralDeCaja.Caption = "Estado General De Caja";
			this.mnuContEstadoGeneralDeCaja.Id = 879;
			this.mnuContEstadoGeneralDeCaja.Name = "mnuContEstadoGeneralDeCaja";
			this.mnuContEstadoGeneralDeCaja.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContEstadoGeneralDeCaja_ItemClick);
			// 
			// MnuRol
			// 
			this.MnuRol.Caption = "&Roles";
			this.MnuRol.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.MnuRol.Id = 136;
			this.MnuRol.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																												new DevExpress.XtraBars.LinkPersistInfo(this.MnuRolPersonal),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolPA),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolPersonalPrueba),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.MnuRolCargos, true),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.MnuRolDpto),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesComisiones, true),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.MnuRolVerif, true),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolFiniquitos),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.MnuRolGenRol, true),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesPagoComisiones),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesDecimoTercerSueldo),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesReportes, true),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.barRolRubro, true),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.MnuRolRepRol),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.MnuRolBloqRol),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.MnuRolIncSuel),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesLiquidacion),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolUtilidad),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.MnuRolAsigCtas, true),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolAnexo),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuTarea_Agr),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesDistribucion, true),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesAdicional, true),
																																												new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesAdicionalFormulas)});
			this.MnuRol.Name = "MnuRol";
			// 
			// MnuRolPersonal
			// 
			this.MnuRolPersonal.Caption = "Personal      ";
			this.MnuRolPersonal.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.MnuRolPersonal.Id = 137;
			this.MnuRolPersonal.ImageIndex = 28;
			this.MnuRolPersonal.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.P));
			this.MnuRolPersonal.Name = "MnuRolPersonal";
			this.MnuRolPersonal.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuRolPersonal_ItemClick);
			// 
			// mnuRolPA
			// 
			this.mnuRolPA.Caption = "Personal Antiguo";
			this.mnuRolPA.Id = 694;
			this.mnuRolPA.Name = "mnuRolPA";
			this.mnuRolPA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolPA_ItemClick);
			// 
			// mnuRolPersonalPrueba
			// 
			this.mnuRolPersonalPrueba.Caption = "Personal a Prueba";
			this.mnuRolPersonalPrueba.Id = 776;
			this.mnuRolPersonalPrueba.Name = "mnuRolPersonalPrueba";
			this.mnuRolPersonalPrueba.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolPersonalPrueba_ItemClick);
			// 
			// MnuRolCargos
			// 
			this.MnuRolCargos.Caption = "Cargos";
			this.MnuRolCargos.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.MnuRolCargos.Id = 138;
			this.MnuRolCargos.ImageIndex = 29;
			this.MnuRolCargos.Name = "MnuRolCargos";
			this.MnuRolCargos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuRolCargos_ItemClick);
			// 
			// MnuRolDpto
			// 
			this.MnuRolDpto.Caption = "Departamento";
			this.MnuRolDpto.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.MnuRolDpto.Id = 139;
			this.MnuRolDpto.ImageIndex = 30;
			this.MnuRolDpto.Name = "MnuRolDpto";
			this.MnuRolDpto.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuRolDpto_ItemClick);
			// 
			// mnuRolesComisiones
			// 
			this.mnuRolesComisiones.Caption = "Comisiones";
			this.mnuRolesComisiones.Id = 691;
			this.mnuRolesComisiones.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolComPIA),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesCom_PA),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuCom_PAGE),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesComisionesPS),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesCom_PL, true),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesCom_PML),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuCom_PMLGE),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesCom_AV, true),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesComAJV),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolComisionesPC, true),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolCom_PorcCarpa),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuCom_PCGE),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesCom_EPV, true),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuCom_Reportes),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolPorCajero)});
			this.mnuRolesComisiones.Name = "mnuRolesComisiones";
			// 
			// mnuRolComPIA
			// 
			this.mnuRolComPIA.Caption = "Presupuesto Inicio de Año ";
			this.mnuRolComPIA.Id = 777;
			this.mnuRolComPIA.Name = "mnuRolComPIA";
			this.mnuRolComPIA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolComPIA_ItemClick);
			// 
			// mnuRolesCom_PA
			// 
			this.mnuRolesCom_PA.Caption = "Presupuesto Anual";
			this.mnuRolesCom_PA.Id = 688;
			this.mnuRolesCom_PA.Name = "mnuRolesCom_PA";
			this.mnuRolesCom_PA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesCom_PA_ItemClick);
			// 
			// mnuCom_PAGE
			// 
			this.mnuCom_PAGE.Caption = "Presupuesto Anual Garantia Extendida";
			this.mnuCom_PAGE.Id = 711;
			this.mnuCom_PAGE.Name = "mnuCom_PAGE";
			this.mnuCom_PAGE.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuCom_PAGE.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCom_PAGE_ItemClick);
			// 
			// mnuRolesComisionesPS
			// 
			this.mnuRolesComisionesPS.Caption = "Presupuesto Semanal";
			this.mnuRolesComisionesPS.Id = 858;
			this.mnuRolesComisionesPS.Name = "mnuRolesComisionesPS";
			this.mnuRolesComisionesPS.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuRolesComisionesPS.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesComisionesPS_ItemClick);
			// 
			// mnuRolesCom_PL
			// 
			this.mnuRolesCom_PL.Caption = "Personal Por Local";
			this.mnuRolesCom_PL.Id = 689;
			this.mnuRolesCom_PL.Name = "mnuRolesCom_PL";
			this.mnuRolesCom_PL.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesCom_PL_ItemClick);
			// 
			// mnuRolesCom_PML
			// 
			this.mnuRolesCom_PML.Caption = "Presupuesto Mensual Por Local";
			this.mnuRolesCom_PML.Id = 690;
			this.mnuRolesCom_PML.Name = "mnuRolesCom_PML";
			this.mnuRolesCom_PML.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuRolesCom_PML.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesCom_PML_ItemClick);
			// 
			// mnuCom_PMLGE
			// 
			this.mnuCom_PMLGE.Caption = "Presupuesto Mensual Por local Garantia Extendida";
			this.mnuCom_PMLGE.Id = 712;
			this.mnuCom_PMLGE.Name = "mnuCom_PMLGE";
			this.mnuCom_PMLGE.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuCom_PMLGE.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCom_PMLGE_ItemClick);
			// 
			// mnuRolesCom_AV
			// 
			this.mnuRolesCom_AV.Caption = "Asignación Mensual de Vendedores";
			this.mnuRolesCom_AV.Id = 679;
			this.mnuRolesCom_AV.Name = "mnuRolesCom_AV";
			this.mnuRolesCom_AV.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesCom_AV_ItemClick);
			// 
			// mnuRolesComAJV
			// 
			this.mnuRolesComAJV.Caption = "Asignación Mensual de Jefes y Supervisores De Ventas";
			this.mnuRolesComAJV.Id = 704;
			this.mnuRolesComAJV.Name = "mnuRolesComAJV";
			this.mnuRolesComAJV.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesComAJV_ItemClick);
			// 
			// mnuRolComisionesPC
			// 
			this.mnuRolComisionesPC.Caption = "Porcentajes de Comisiones";
			this.mnuRolComisionesPC.Id = 699;
			this.mnuRolComisionesPC.Name = "mnuRolComisionesPC";
			this.mnuRolComisionesPC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolComisionesPC_ItemClick);
			// 
			// mnuRolCom_PorcCarpa
			// 
			this.mnuRolCom_PorcCarpa.Caption = "Porcentajes de Carpa";
			this.mnuRolCom_PorcCarpa.Id = 796;
			this.mnuRolCom_PorcCarpa.Name = "mnuRolCom_PorcCarpa";
			this.mnuRolCom_PorcCarpa.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolCom_PorcCarpa_ItemClick);
			// 
			// mnuCom_PCGE
			// 
			this.mnuCom_PCGE.Caption = "Porcentajes de Comisiones Garantia Extendida";
			this.mnuCom_PCGE.Id = 713;
			this.mnuCom_PCGE.Name = "mnuCom_PCGE";
			this.mnuCom_PCGE.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCom_PCGE_ItemClick);
			// 
			// mnuRolesCom_EPV
			// 
			this.mnuRolesCom_EPV.Caption = "Evaluación Personal De Ventas";
			this.mnuRolesCom_EPV.Id = 775;
			this.mnuRolesCom_EPV.Name = "mnuRolesCom_EPV";
			this.mnuRolesCom_EPV.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesCom_EPV_ItemClick);
			// 
			// mnuCom_Reportes
			// 
			this.mnuCom_Reportes.Caption = "Reportes";
			this.mnuCom_Reportes.Id = 715;
			this.mnuCom_Reportes.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCom_ReportesRM),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCom_ReporteLM)});
			this.mnuCom_Reportes.Name = "mnuCom_Reportes";
			// 
			// mnuCom_ReportesRM
			// 
			this.mnuCom_ReportesRM.Caption = "Resumen Mensual";
			this.mnuCom_ReportesRM.Id = 716;
			this.mnuCom_ReportesRM.Name = "mnuCom_ReportesRM";
			this.mnuCom_ReportesRM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCom_ReportesRM_ItemClick);
			// 
			// mnuCom_ReporteLM
			// 
			this.mnuCom_ReporteLM.Caption = "Movimientos Comisiones";
			this.mnuCom_ReporteLM.Id = 974;
			this.mnuCom_ReporteLM.Name = "mnuCom_ReporteLM";
			this.mnuCom_ReporteLM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCom_ReporteLM_ItemClick);
			// 
			// mnuRolPorCajero
			// 
			this.mnuRolPorCajero.Caption = "Porcentajes Cajero";
			this.mnuRolPorCajero.Id = 993;
			this.mnuRolPorCajero.Name = "mnuRolPorCajero";
			this.mnuRolPorCajero.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolPorCajero_ItemClick);
			// 
			// MnuRolVerif
			// 
			this.MnuRolVerif.Caption = "Anticipos Empleados";
			this.MnuRolVerif.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.MnuRolVerif.Id = 144;
			this.MnuRolVerif.Name = "MnuRolVerif";
			this.MnuRolVerif.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuRolVerif_ItemClick);
			// 
			// mnuRolFiniquitos
			// 
			this.mnuRolFiniquitos.Caption = "Finiquitos";
			this.mnuRolFiniquitos.Id = 550;
			this.mnuRolFiniquitos.Name = "mnuRolFiniquitos";
			this.mnuRolFiniquitos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolFiniquitos_ItemClick);
			// 
			// MnuRolGenRol
			// 
			this.MnuRolGenRol.Caption = "Generar Roles";
			this.MnuRolGenRol.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.MnuRolGenRol.Id = 140;
			this.MnuRolGenRol.ImageIndex = 34;
			this.MnuRolGenRol.Name = "MnuRolGenRol";
			this.MnuRolGenRol.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuRolGenRol_ItemClick);
			// 
			// mnuRolesPagoComisiones
			// 
			this.mnuRolesPagoComisiones.Caption = "Generar Comisiones";
			this.mnuRolesPagoComisiones.Id = 703;
			this.mnuRolesPagoComisiones.Name = "mnuRolesPagoComisiones";
			this.mnuRolesPagoComisiones.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesPagoComisiones_ItemClick);
			// 
			// mnuRolesDecimoTercerSueldo
			// 
			this.mnuRolesDecimoTercerSueldo.Caption = "Genera Decimo Tercer Sueldo";
			this.mnuRolesDecimoTercerSueldo.Id = 887;
			this.mnuRolesDecimoTercerSueldo.Name = "mnuRolesDecimoTercerSueldo";
			this.mnuRolesDecimoTercerSueldo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesDecimoTercerSueldo_ItemClick);
			// 
			// mnuRolesReportes
			// 
			this.mnuRolesReportes.Caption = "Reportes";
			this.mnuRolesReportes.Id = 708;
			this.mnuRolesReportes.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesReportesPVF),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesReportesVXM),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesReportesUniformes),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesReportesNominaGeneral),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolReportesAE),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolesReportesHHT),
																																																	new DevExpress.XtraBars.LinkPersistInfo(this.mnuRolReportesIND)});
			this.mnuRolesReportes.Name = "mnuRolesReportes";
			// 
			// mnuRolesReportesPVF
			// 
			this.mnuRolesReportesPVF.Caption = "Permisos - Vacaciones - Faltas";
			this.mnuRolesReportesPVF.Id = 709;
			this.mnuRolesReportesPVF.Name = "mnuRolesReportesPVF";
			this.mnuRolesReportesPVF.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesReportesPVF_ItemClick);
			// 
			// mnuRolesReportesVXM
			// 
			this.mnuRolesReportesVXM.Caption = "Vendedores Por Mes";
			this.mnuRolesReportesVXM.Id = 719;
			this.mnuRolesReportesVXM.Name = "mnuRolesReportesVXM";
			this.mnuRolesReportesVXM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesReportesVXM_ItemClick);
			// 
			// mnuRolesReportesUniformes
			// 
			this.mnuRolesReportesUniformes.Caption = "Uniformes";
			this.mnuRolesReportesUniformes.Id = 749;
			this.mnuRolesReportesUniformes.Name = "mnuRolesReportesUniformes";
			this.mnuRolesReportesUniformes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesReportesUniformes_ItemClick);
			// 
			// mnuRolesReportesNominaGeneral
			// 
			this.mnuRolesReportesNominaGeneral.Caption = "Nomina General";
			this.mnuRolesReportesNominaGeneral.Id = 770;
			this.mnuRolesReportesNominaGeneral.Name = "mnuRolesReportesNominaGeneral";
			this.mnuRolesReportesNominaGeneral.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesReportesNominaGeneral_ItemClick);
			// 
			// mnuRolReportesAE
			// 
			this.mnuRolReportesAE.Caption = "Anticipos Empleados";
			this.mnuRolReportesAE.Id = 831;
			this.mnuRolReportesAE.Name = "mnuRolReportesAE";
			this.mnuRolReportesAE.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolReportesAE_ItemClick);
			// 
			// mnuRolesReportesHHT
			// 
			this.mnuRolesReportesHHT.Caption = "Horas Hombre Trabajadas";
			this.mnuRolesReportesHHT.Id = 837;
			this.mnuRolesReportesHHT.Name = "mnuRolesReportesHHT";
			this.mnuRolesReportesHHT.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesReportesHHT_ItemClick);
			// 
			// mnuRolReportesIND
			// 
			this.mnuRolReportesIND.Caption = "Ingresos Por Departamento";
			this.mnuRolReportesIND.Id = 866;
			this.mnuRolReportesIND.Name = "mnuRolReportesIND";
			this.mnuRolReportesIND.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolReportesIND_ItemClick);
			// 
			// barRolRubro
			// 
			this.barRolRubro.Caption = "Rubros";
			this.barRolRubro.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.barRolRubro.Id = 240;
			this.barRolRubro.Name = "barRolRubro";
			this.barRolRubro.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barRolRubro_ItemClick);
			// 
			// MnuRolRepRol
			// 
			this.MnuRolRepRol.Caption = "Reporte de Roles ";
			this.MnuRolRepRol.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.MnuRolRepRol.Id = 142;
			this.MnuRolRepRol.Name = "MnuRolRepRol";
			this.MnuRolRepRol.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuRolRepRol_ItemClick);
			// 
			// MnuRolBloqRol
			// 
			this.MnuRolBloqRol.Caption = "Bloquear Registro de Roles";
			this.MnuRolBloqRol.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.MnuRolBloqRol.Id = 145;
			this.MnuRolBloqRol.ImageIndex = 31;
			this.MnuRolBloqRol.Name = "MnuRolBloqRol";
			this.MnuRolBloqRol.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuRolBloqRol_ItemClick);
			// 
			// MnuRolIncSuel
			// 
			this.MnuRolIncSuel.Caption = "Incrementar Sueldos";
			this.MnuRolIncSuel.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.MnuRolIncSuel.Id = 143;
			this.MnuRolIncSuel.ImageIndex = 32;
			this.MnuRolIncSuel.Name = "MnuRolIncSuel";
			this.MnuRolIncSuel.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuRolIncSuel_ItemClick);
			// 
			// mnuRolesLiquidacion
			// 
			this.mnuRolesLiquidacion.Caption = "Liquidaciones";
			this.mnuRolesLiquidacion.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.mnuRolesLiquidacion.Id = 280;
			this.mnuRolesLiquidacion.Name = "mnuRolesLiquidacion";
			this.mnuRolesLiquidacion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesLiquidacion_ItemClick);
			// 
			// mnuRolUtilidad
			// 
			this.mnuRolUtilidad.Caption = "Repartición de Utilidades";
			this.mnuRolUtilidad.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.mnuRolUtilidad.Id = 296;
			this.mnuRolUtilidad.Name = "mnuRolUtilidad";
			this.mnuRolUtilidad.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolUtilidad_ItemClick);
			// 
			// MnuRolAsigCtas
			// 
			this.MnuRolAsigCtas.Caption = "Impuesto a la Renta";
			this.MnuRolAsigCtas.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.MnuRolAsigCtas.Id = 146;
			this.MnuRolAsigCtas.Name = "MnuRolAsigCtas";
			this.MnuRolAsigCtas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuRolAsigCtas_ItemClick);
			// 
			// mnuRolAnexo
			// 
			this.mnuRolAnexo.Caption = "Anexo Renta RDEP";
			this.mnuRolAnexo.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.mnuRolAnexo.Id = 285;
			this.mnuRolAnexo.Name = "mnuRolAnexo";
			this.mnuRolAnexo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolAnexo_ItemClick);
			// 
			// mnuTarea_Agr
			// 
			this.mnuTarea_Agr.Caption = "Tareas Agrícolas";
			this.mnuTarea_Agr.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.mnuTarea_Agr.Id = 308;
			this.mnuTarea_Agr.Name = "mnuTarea_Agr";
			this.mnuTarea_Agr.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuProdAgro_ItemClick);
			// 
			// mnuRolesDistribucion
			// 
			this.mnuRolesDistribucion.Caption = "Distribución Centros Costos";
			this.mnuRolesDistribucion.Id = 406;
			this.mnuRolesDistribucion.Name = "mnuRolesDistribucion";
			this.mnuRolesDistribucion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesDistribucion_ItemClick);
			// 
			// mnuRolesAdicional
			// 
			this.mnuRolesAdicional.Caption = "Adicionales";
			this.mnuRolesAdicional.Id = 422;
			this.mnuRolesAdicional.Name = "mnuRolesAdicional";
			this.mnuRolesAdicional.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesAdicional_ItemClick);
			// 
			// mnuRolesAdicionalFormulas
			// 
			this.mnuRolesAdicionalFormulas.Caption = "Adicionales Fórmulas";
			this.mnuRolesAdicionalFormulas.Id = 425;
			this.mnuRolesAdicionalFormulas.Name = "mnuRolesAdicionalFormulas";
			this.mnuRolesAdicionalFormulas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRolesAdicionalFormulas_ItemClick);
			// 
			// MnuCompras
			// 
			this.MnuCompras.Caption = "Co&mpras";
			this.MnuCompras.Id = 525;
			this.MnuCompras.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactProvee),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactOrdComp, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuCompraAODC, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactCompras, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactDevCompras, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactNotaCC, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactLotimp, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuComprasLAC, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuFactVentasConsignacion, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuContAntProv, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuContASF, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuContCXP, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuContResCXP, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuComprasGDA, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactArtic),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCompraCSA),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuVentasReportesVA, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuComprasCatalogo, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactCombos, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnnuFactLotesPromociones, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuComprasPromociones),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuFactActPrecio, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactLotesPromVen, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuComprasRepSpifVend),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuComprasAgenda, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuCompraGestion, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuComprasLiquidaGastos, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuComprasComboZOOM, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuMatrizDsctos, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuComprasReportes),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuComprasProformas),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuComprasPromoFact)});
			this.MnuCompras.Name = "MnuCompras";
			// 
			// MnuFactProvee
			// 
			this.MnuFactProvee.Caption = "Proveedores";
			this.MnuFactProvee.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactProvee.Id = 127;
			this.MnuFactProvee.ImageIndex = 27;
			this.MnuFactProvee.Name = "MnuFactProvee";
			this.MnuFactProvee.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactProvee_ItemClick);
			// 
			// MnuFactOrdComp
			// 
			this.MnuFactOrdComp.Caption = "Orden de Compra";
			this.MnuFactOrdComp.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactOrdComp.Id = 106;
			this.MnuFactOrdComp.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O));
			this.MnuFactOrdComp.Name = "MnuFactOrdComp";
			this.MnuFactOrdComp.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactOrdComp_ItemClick);
			// 
			// mnuCompraAODC
			// 
			this.mnuCompraAODC.Caption = "Aprobación Ordenes De compra";
			this.mnuCompraAODC.Id = 968;
			this.mnuCompraAODC.Name = "mnuCompraAODC";
			this.mnuCompraAODC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCompraAODC_ItemClick);
			// 
			// MnuFactCompras
			// 
			this.MnuFactCompras.Caption = "Compras";
			this.MnuFactCompras.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactCompras.Id = 105;
			this.MnuFactCompras.ImageIndex = 18;
			this.MnuFactCompras.ItemShortcut = new DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F10);
			this.MnuFactCompras.Name = "MnuFactCompras";
			this.MnuFactCompras.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactCompras_ItemClick);
			// 
			// MnuFactDevCompras
			// 
			this.MnuFactDevCompras.Caption = "Notas de Crédito";
			this.MnuFactDevCompras.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactDevCompras.Id = 110;
			this.MnuFactDevCompras.Name = "MnuFactDevCompras";
			this.MnuFactDevCompras.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactDevCompras_ItemClick);
			// 
			// MnuFactNotaCC
			// 
			this.MnuFactNotaCC.Caption = "Nota de Débito Compra";
			this.MnuFactNotaCC.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactNotaCC.Id = 257;
			this.MnuFactNotaCC.Name = "MnuFactNotaCC";
			this.MnuFactNotaCC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactNotaCC_ItemClick);
			// 
			// MnuFactLotimp
			// 
			this.MnuFactLotimp.Caption = "Liquidación de Importaciones";
			this.MnuFactLotimp.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactLotimp.Id = 242;
			this.MnuFactLotimp.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.I));
			this.MnuFactLotimp.Name = "MnuFactLotimp";
			this.MnuFactLotimp.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactImportacion_ItemClick);
			// 
			// mnuComprasLAC
			// 
			this.mnuComprasLAC.Caption = "Lista Articulos En Consignacion";
			this.mnuComprasLAC.Id = 976;
			this.mnuComprasLAC.Name = "mnuComprasLAC";
			this.mnuComprasLAC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasLAC_ItemClick);
			// 
			// mnuFactVentasConsignacion
			// 
			this.mnuFactVentasConsignacion.Caption = "Venta de artículos en consignación";
			this.mnuFactVentasConsignacion.Id = 519;
			this.mnuFactVentasConsignacion.Name = "mnuFactVentasConsignacion";
			this.mnuFactVentasConsignacion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFactVentasConsignacion_ItemClick);
			// 
			// mnuContAntProv
			// 
			this.mnuContAntProv.Caption = "Anticipos a Proveedor";
			this.mnuContAntProv.Id = 561;
			this.mnuContAntProv.Name = "mnuContAntProv";
			this.mnuContAntProv.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContAntProv_ItemClick);
			// 
			// mnuContASF
			// 
			this.mnuContASF.Caption = "Anticipos Sin Factura";
			this.mnuContASF.Id = 564;
			this.mnuContASF.Name = "mnuContASF";
			this.mnuContASF.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContASF_ItemClick);
			// 
			// MnuContCXP
			// 
			this.MnuContCXP.Caption = "Cuentas por Pagar";
			this.MnuContCXP.Id = 517;
			this.MnuContCXP.Name = "MnuContCXP";
			this.MnuContCXP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContCXP_ItemClick);
			// 
			// mnuContResCXP
			// 
			this.mnuContResCXP.Caption = "Resumen de Cuentas Por Pagar";
			this.mnuContResCXP.Id = 547;
			this.mnuContResCXP.Name = "mnuContResCXP";
			this.mnuContResCXP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContResCXP_ItemClick);
			// 
			// mnuComprasGDA
			// 
			this.mnuComprasGDA.Caption = "Grupos De Articulos";
			this.mnuComprasGDA.Id = 668;
			this.mnuComprasGDA.Name = "mnuComprasGDA";
			this.mnuComprasGDA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasGDA_ItemClick);
			// 
			// MnuFactArtic
			// 
			this.MnuFactArtic.Caption = "Artículos   ";
			this.MnuFactArtic.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactArtic.Id = 125;
			this.MnuFactArtic.ImageIndex = 26;
			this.MnuFactArtic.Name = "MnuFactArtic";
			this.MnuFactArtic.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactArtic_ItemClick);
			// 
			// mnuCompraCSA
			// 
			this.mnuCompraCSA.Caption = "Caracteristicas por Subgrupo de Articulos";
			this.mnuCompraCSA.Id = 784;
			this.mnuCompraCSA.Name = "mnuCompraCSA";
			this.mnuCompraCSA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCompraCSA_ItemClick);
			// 
			// mnuVentasReportesVA
			// 
			this.mnuVentasReportesVA.Caption = "Ventas Por Articulos";
			this.mnuVentasReportesVA.Id = 804;
			this.mnuVentasReportesVA.Name = "mnuVentasReportesVA";
			this.mnuVentasReportesVA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesVA_ItemClick);
			// 
			// mnuComprasCatalogo
			// 
			this.mnuComprasCatalogo.Caption = "Catalogo";
			this.mnuComprasCatalogo.Id = 852;
			this.mnuComprasCatalogo.Name = "mnuComprasCatalogo";
			this.mnuComprasCatalogo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasCatalogo_ItemClick);
			// 
			// MnuFactCombos
			// 
			this.MnuFactCombos.Caption = "Combos";
			this.MnuFactCombos.Id = 467;
			this.MnuFactCombos.Name = "MnuFactCombos";
			this.MnuFactCombos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactCombos_ItemClick);
			// 
			// MnnuFactLotesPromociones
			// 
			this.MnnuFactLotesPromociones.Caption = "Promociones";
			this.MnnuFactLotesPromociones.Id = 468;
			this.MnnuFactLotesPromociones.Name = "MnnuFactLotesPromociones";
			this.MnnuFactLotesPromociones.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnnuFactLotesPromociones_ItemClick);
			// 
			// mnuComprasPromociones
			// 
			this.mnuComprasPromociones.Caption = "Registro de Promociones Activas";
			this.mnuComprasPromociones.Id = 602;
			this.mnuComprasPromociones.Name = "mnuComprasPromociones";
			this.mnuComprasPromociones.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasPromociones_ItemClick);
			// 
			// mnuFactActPrecio
			// 
			this.mnuFactActPrecio.Caption = "Actualización de Precios";
			this.mnuFactActPrecio.Id = 512;
			this.mnuFactActPrecio.Name = "mnuFactActPrecio";
			this.mnuFactActPrecio.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFactActPrecio_ItemClick);
			// 
			// MnuFactLotesPromVen
			// 
			this.MnuFactLotesPromVen.Caption = "SPIF para Vendedores";
			this.MnuFactLotesPromVen.Id = 487;
			this.MnuFactLotesPromVen.Name = "MnuFactLotesPromVen";
			this.MnuFactLotesPromVen.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactLotesPromVen_ItemClick);
			// 
			// mnuComprasRepSpifVend
			// 
			this.mnuComprasRepSpifVend.Caption = "Reporte SPIF para Vendedores";
			this.mnuComprasRepSpifVend.Id = 570;
			this.mnuComprasRepSpifVend.Name = "mnuComprasRepSpifVend";
			this.mnuComprasRepSpifVend.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasRepSpifVend_ItemClick);
			// 
			// mnuComprasAgenda
			// 
			this.mnuComprasAgenda.Caption = "Agenda";
			this.mnuComprasAgenda.Id = 543;
			this.mnuComprasAgenda.Name = "mnuComprasAgenda";
			this.mnuComprasAgenda.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasAgenda_ItemClick);
			// 
			// mnuCompraGestion
			// 
			this.mnuCompraGestion.Caption = "Gestión Cliente";
			this.mnuCompraGestion.Id = 544;
			this.mnuCompraGestion.Name = "mnuCompraGestion";
			this.mnuCompraGestion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCompraGestion_ItemClick);
			// 
			// mnuComprasLiquidaGastos
			// 
			this.mnuComprasLiquidaGastos.Caption = "Liquidación de gastos";
			this.mnuComprasLiquidaGastos.Id = 541;
			this.mnuComprasLiquidaGastos.Name = "mnuComprasLiquidaGastos";
			this.mnuComprasLiquidaGastos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasLiquidaGastos_ItemClick);
			// 
			// mnuComprasComboZOOM
			// 
			this.mnuComprasComboZOOM.Caption = "Combos ZOOM";
			this.mnuComprasComboZOOM.Id = 539;
			this.mnuComprasComboZOOM.Name = "mnuComprasComboZOOM";
			this.mnuComprasComboZOOM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuComprasComboZOOM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasComboZOOM_ItemClick);
			// 
			// mnuMatrizDsctos
			// 
			this.mnuMatrizDsctos.Caption = "Matriz de Descuentos";
			this.mnuMatrizDsctos.Id = 542;
			this.mnuMatrizDsctos.Name = "mnuMatrizDsctos";
			this.mnuMatrizDsctos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuMatrizDsctos_ItemClick);
			// 
			// mnuComprasReportes
			// 
			this.mnuComprasReportes.Caption = "Reportes";
			this.mnuComprasReportes.Id = 587;
			this.mnuComprasReportes.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuCompraReportesCPA),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuComprasAPC),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuComprasRepFP),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuComprasRepCodigoArticulos)});
			this.mnuComprasReportes.Name = "mnuComprasReportes";
			// 
			// mnuCompraReportesCPA
			// 
			this.mnuCompraReportesCPA.Caption = "Compras Por Articulo";
			this.mnuCompraReportesCPA.Id = 599;
			this.mnuCompraReportesCPA.Name = "mnuCompraReportesCPA";
			this.mnuCompraReportesCPA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCompraReportesCPA_ItemClick);
			// 
			// mnuComprasAPC
			// 
			this.mnuComprasAPC.Caption = "Articulos Por Comprar";
			this.mnuComprasAPC.Id = 706;
			this.mnuComprasAPC.Name = "mnuComprasAPC";
			this.mnuComprasAPC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasAPC_ItemClick);
			// 
			// mnuComprasRepFP
			// 
			this.mnuComprasRepFP.Caption = "Flujo Proveedores";
			this.mnuComprasRepFP.Id = 676;
			this.mnuComprasRepFP.Name = "mnuComprasRepFP";
			this.mnuComprasRepFP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasRepFP_ItemClick);
			// 
			// mnuComprasRepCodigoArticulos
			// 
			this.mnuComprasRepCodigoArticulos.Caption = "Codigo Articulos";
			this.mnuComprasRepCodigoArticulos.Id = 848;
			this.mnuComprasRepCodigoArticulos.Name = "mnuComprasRepCodigoArticulos";
			this.mnuComprasRepCodigoArticulos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasRepCodigoArticulos_ItemClick);
			// 
			// mnuComprasProformas
			// 
			this.mnuComprasProformas.Caption = "Proforma";
			this.mnuComprasProformas.Id = 865;
			this.mnuComprasProformas.Name = "mnuComprasProformas";
			this.mnuComprasProformas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasProformas_ItemClick);
			// 
			// mnuComprasPromoFact
			// 
			this.mnuComprasPromoFact.Caption = "Promociones Factura";
			this.mnuComprasPromoFact.Id = 990;
			this.mnuComprasPromoFact.Name = "mnuComprasPromoFact";
			this.mnuComprasPromoFact.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuComprasPromoFact_ItemClick);
			// 
			// mnuVentas
			// 
			this.mnuVentas.Caption = "&Ventas";
			this.mnuVentas.Id = 527;
			this.mnuVentas.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactClientes),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactVentas, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.barButtonItem66, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactPreFactura, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													 new DevExpress.XtraBars.LinkPersistInfo(((DevExpress.XtraBars.BarLinkUserDefines)((DevExpress.XtraBars.BarLinkUserDefines.Caption | DevExpress.XtraBars.BarLinkUserDefines.PaintStyle))), this.mnuVentasPFR, "Prefactura de Remate", false, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactPedClientes, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactReservClientes, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuVentasCotizacion, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactConsignacion, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuVentaLiqCons, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasECW),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactDevVentas, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Standard),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactNotaCV, DevExpress.XtraBars.BarItemPaintStyle.Standard),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.MnuFactAnticipos, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Standard),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasDA),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasCNOS, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasRPA),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentaAutDistribuidor, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasCB),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContCuadreCaja, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContDepositoVouchers),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContMantTarj),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasResumenDiario, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasResumenComisiones),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasMarcador),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasPlanifVentas, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasPlanifJefeVentas),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasPC),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuContCXC, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasImpresionDePrecios, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasIPP),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasIPR),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasARP, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasRP),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.barSubItem3, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentaPromocionesActivas, true)});
			this.mnuVentas.Name = "mnuVentas";
			// 
			// MnuFactClientes
			// 
			this.MnuFactClientes.Caption = "Clientes";
			this.MnuFactClientes.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactClientes.Id = 126;
			this.MnuFactClientes.ImageIndex = 21;
			this.MnuFactClientes.Name = "MnuFactClientes";
			this.MnuFactClientes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactClientes_ItemClick);
			// 
			// MnuFactVentas
			// 
			this.MnuFactVentas.Caption = "Ventas         ";
			this.MnuFactVentas.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactVentas.Id = 107;
			this.MnuFactVentas.ImageIndex = 22;
			this.MnuFactVentas.ItemShortcut = new DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F6);
			this.MnuFactVentas.Name = "MnuFactVentas";
			this.MnuFactVentas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactVentas_ItemClick);
			// 
			// barButtonItem66
			// 
			this.barButtonItem66.Caption = "Ventas de Remate";
			this.barButtonItem66.Id = 763;
			this.barButtonItem66.Name = "barButtonItem66";
			this.barButtonItem66.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem66_ItemClick);
			// 
			// MnuFactPreFactura
			// 
			this.MnuFactPreFactura.Caption = "Prefactura";
			this.MnuFactPreFactura.Id = 458;
			this.MnuFactPreFactura.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F6));
			this.MnuFactPreFactura.Name = "MnuFactPreFactura";
			this.MnuFactPreFactura.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactPreFactura_ItemClick);
			// 
			// mnuVentasPFR
			// 
			this.mnuVentasPFR.Caption = "Prefactura de Remate";
			this.mnuVentasPFR.Id = 822;
			this.mnuVentasPFR.Name = "mnuVentasPFR";
			this.mnuVentasPFR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasPFR_ItemClick);
			// 
			// MnuFactPedClientes
			// 
			this.MnuFactPedClientes.Caption = "Pedidos";
			this.MnuFactPedClientes.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactPedClientes.Id = 116;
			this.MnuFactPedClientes.Name = "MnuFactPedClientes";
			this.MnuFactPedClientes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactPedClientes_ItemClick);
			// 
			// MnuFactReservClientes
			// 
			this.MnuFactReservClientes.Caption = "Reserva";
			this.MnuFactReservClientes.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactReservClientes.Id = 113;
			this.MnuFactReservClientes.Name = "MnuFactReservClientes";
			this.MnuFactReservClientes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactReservClientes_ItemClick);
			// 
			// mnuVentasCotizacion
			// 
			this.mnuVentasCotizacion.Caption = "Cotización";
			this.mnuVentasCotizacion.Id = 538;
			this.mnuVentasCotizacion.Name = "mnuVentasCotizacion";
			this.mnuVentasCotizacion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasCotizacion_ItemClick);
			// 
			// MnuFactConsignacion
			// 
			this.MnuFactConsignacion.Caption = "Consignación";
			this.MnuFactConsignacion.Id = 503;
			this.MnuFactConsignacion.Name = "MnuFactConsignacion";
			this.MnuFactConsignacion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactConsignacion_ItemClick);
			// 
			// mnuVentaLiqCons
			// 
			this.mnuVentaLiqCons.Caption = "Liquidación de Consignación";
			this.mnuVentaLiqCons.Id = 536;
			this.mnuVentaLiqCons.Name = "mnuVentaLiqCons";
			this.mnuVentaLiqCons.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentaLiqCons_ItemClick);
			// 
			// mnuVentasECW
			// 
			this.mnuVentasECW.Caption = "Codigos WEB";
			this.mnuVentasECW.Id = 841;
			this.mnuVentasECW.Name = "mnuVentasECW";
			this.mnuVentasECW.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasECW_ItemClick);
			// 
			// MnuFactDevVentas
			// 
			this.MnuFactDevVentas.Caption = "Notas de Crédito";
			this.MnuFactDevVentas.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactDevVentas.Id = 109;
			this.MnuFactDevVentas.Name = "MnuFactDevVentas";
			this.MnuFactDevVentas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactDevVentas_ItemClick);
			// 
			// MnuFactNotaCV
			// 
			this.MnuFactNotaCV.Caption = "Notas de Débito";
			this.MnuFactNotaCV.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactNotaCV.Id = 256;
			this.MnuFactNotaCV.Name = "MnuFactNotaCV";
			this.MnuFactNotaCV.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactNotaCV_ItemClick);
			// 
			// MnuFactAnticipos
			// 
			this.MnuFactAnticipos.Caption = "Anticipos";
			this.MnuFactAnticipos.Id = 492;
			this.MnuFactAnticipos.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F11));
			this.MnuFactAnticipos.Name = "MnuFactAnticipos";
			this.MnuFactAnticipos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactAnticipos_ItemClick);
			// 
			// mnuVentasDA
			// 
			this.mnuVentasDA.Caption = "Devolución Anticipos";
			this.mnuVentasDA.Id = 573;
			this.mnuVentasDA.Name = "mnuVentasDA";
			this.mnuVentasDA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasDA_ItemClick);
			// 
			// mnuVentasCNOS
			// 
			this.mnuVentasCNOS.Caption = "Consulta Numeros de Operación Solidario";
			this.mnuVentasCNOS.Id = 862;
			this.mnuVentasCNOS.Name = "mnuVentasCNOS";
			this.mnuVentasCNOS.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasCNOS_ItemClick);
			// 
			// mnuVentasRPA
			// 
			this.mnuVentasRPA.Caption = "Cobros Pago Agil";
			this.mnuVentasRPA.Id = 857;
			this.mnuVentasRPA.Name = "mnuVentasRPA";
			this.mnuVentasRPA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasRPA_ItemClick);
			// 
			// mnuVentaAutDistribuidor
			// 
			this.mnuVentaAutDistribuidor.Caption = "Autorización Distribuidor";
			this.mnuVentaAutDistribuidor.Id = 537;
			this.mnuVentaAutDistribuidor.Name = "mnuVentaAutDistribuidor";
			this.mnuVentaAutDistribuidor.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentaAutDistribuidor_ItemClick);
			// 
			// mnuVentasCB
			// 
			this.mnuVentasCB.Caption = "Codigos de Barras";
			this.mnuVentasCB.Id = 535;
			this.mnuVentasCB.Name = "mnuVentasCB";
			this.mnuVentasCB.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasCB_ItemClick);
			// 
			// MnuContCuadreCaja
			// 
			this.MnuContCuadreCaja.Caption = "Cuadre de Caja";
			this.MnuContCuadreCaja.Id = 488;
			this.MnuContCuadreCaja.ItemShortcut = new DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F2);
			this.MnuContCuadreCaja.Name = "MnuContCuadreCaja";
			this.MnuContCuadreCaja.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContCuadreCaja_ItemClick);
			// 
			// MnuContDepositoVouchers
			// 
			this.MnuContDepositoVouchers.Caption = "Deposito de Vouchers";
			this.MnuContDepositoVouchers.Id = 516;
			this.MnuContDepositoVouchers.Name = "MnuContDepositoVouchers";
			this.MnuContDepositoVouchers.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContDepositoVouchers_ItemClick);
			// 
			// MnuContMantTarj
			// 
			this.MnuContMantTarj.Caption = "Mantenimiento de Tarjetas";
			this.MnuContMantTarj.Id = 509;
			this.MnuContMantTarj.Name = "MnuContMantTarj";
			this.MnuContMantTarj.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContMantTarj_ItemClick);
			// 
			// mnuVentasResumenDiario
			// 
			this.mnuVentasResumenDiario.Caption = "Resumen Diario de Ventas";
			this.mnuVentasResumenDiario.Id = 533;
			this.mnuVentasResumenDiario.Name = "mnuVentasResumenDiario";
			this.mnuVentasResumenDiario.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasResumenDiario_ItemClick);
			// 
			// mnuVentasResumenComisiones
			// 
			this.mnuVentasResumenComisiones.Caption = "Resumen De Comisiones ";
			this.mnuVentasResumenComisiones.Id = 701;
			this.mnuVentasResumenComisiones.Name = "mnuVentasResumenComisiones";
			this.mnuVentasResumenComisiones.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasResumenComisiones_ItemClick);
			// 
			// mnuVentasMarcador
			// 
			this.mnuVentasMarcador.Caption = "Marcador";
			this.mnuVentasMarcador.Id = 720;
			this.mnuVentasMarcador.Name = "mnuVentasMarcador";
			this.mnuVentasMarcador.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasMarcador_ItemClick);
			// 
			// mnuVentasPlanifVentas
			// 
			this.mnuVentasPlanifVentas.Caption = "Planificación Mensual de Ventas";
			this.mnuVentasPlanifVentas.Id = 698;
			this.mnuVentasPlanifVentas.Name = "mnuVentasPlanifVentas";
			this.mnuVentasPlanifVentas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasPlanifVentas_ItemClick);
			// 
			// mnuVentasPlanifJefeVentas
			// 
			this.mnuVentasPlanifJefeVentas.Caption = "Planificación Mensual Jefes De Ventas";
			this.mnuVentasPlanifJefeVentas.Id = 767;
			this.mnuVentasPlanifJefeVentas.Name = "mnuVentasPlanifJefeVentas";
			this.mnuVentasPlanifJefeVentas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasPlanifJefeVentas_ItemClick);
			// 
			// mnuVentasPC
			// 
			this.mnuVentasPC.Caption = "Planificación Semanal de Carpa";
			this.mnuVentasPC.Id = 692;
			this.mnuVentasPC.Name = "mnuVentasPC";
			this.mnuVentasPC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasPC_ItemClick);
			// 
			// MnuContCXC
			// 
			this.MnuContCXC.Caption = "Cuentas por Cobrar";
			this.MnuContCXC.Id = 518;
			this.MnuContCXC.Name = "MnuContCXC";
			this.MnuContCXC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContCXC_ItemClick);
			// 
			// mnuVentasImpresionDePrecios
			// 
			this.mnuVentasImpresionDePrecios.Caption = "Impresión de Precios";
			this.mnuVentasImpresionDePrecios.Id = 584;
			this.mnuVentasImpresionDePrecios.Name = "mnuVentasImpresionDePrecios";
			this.mnuVentasImpresionDePrecios.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasImpresionDePrecios_ItemClick);
			// 
			// mnuVentasIPP
			// 
			this.mnuVentasIPP.Caption = "Impresión de Precios Promociones";
			this.mnuVentasIPP.Id = 612;
			this.mnuVentasIPP.Name = "mnuVentasIPP";
			this.mnuVentasIPP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasIPP_ItemClick);
			// 
			// mnuVentasIPR
			// 
			this.mnuVentasIPR.Caption = "Impresión de Precios De Remate";
			this.mnuVentasIPR.Id = 798;
			this.mnuVentasIPR.Name = "mnuVentasIPR";
			this.mnuVentasIPR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasIPR_ItemClick);
			// 
			// mnuVentasARP
			// 
			this.mnuVentasARP.Caption = "Actas de Recepción de Productos";
			this.mnuVentasARP.Id = 673;
			this.mnuVentasARP.Name = "mnuVentasARP";
			this.mnuVentasARP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasARP_ItemClick);
			// 
			// mnuVentasRP
			// 
			this.mnuVentasRP.Caption = "Recepción de Productos";
			this.mnuVentasRP.Id = 840;
			this.mnuVentasRP.Name = "mnuVentasRP";
			this.mnuVentasRP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasRP_ItemClick);
			// 
			// barSubItem3
			// 
			this.barSubItem3.Caption = "Reportes";
			this.barSubItem3.Id = 558;
			this.barSubItem3.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReporteCompara),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesPorMarca, true),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasRVML),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasRepMU, true),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentaRMU),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentaRMUA),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesAIM),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesRF, true),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesVPD),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesFFP),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasRepGE),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasRepRVF),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesComparativos),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesVentasPorMes),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuReporteVentasMensualesPorLocal),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesTC),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesDTC),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesEncuestas, true),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesCXC),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuContResCXC),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasReportesVPA),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasComisiones),
																																														 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuVentasReportesVentasPorArticulos, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																														 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuVentasReportesVentasarticuloPorMes, DevExpress.XtraBars.BarItemPaintStyle.Caption)});
			this.barSubItem3.Name = "barSubItem3";
			// 
			// mnuVentasReporteCompara
			// 
			this.mnuVentasReporteCompara.Caption = "Comparación De Ventas Easy Max Latium";
			this.mnuVentasReporteCompara.Id = 559;
			this.mnuVentasReporteCompara.Name = "mnuVentasReporteCompara";
			this.mnuVentasReporteCompara.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuVentasReporteCompara.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReporteCompara_ItemClick);
			// 
			// mnuVentasReportesPorMarca
			// 
			this.mnuVentasReportesPorMarca.Caption = "Ventas Por Marca";
			this.mnuVentasReportesPorMarca.Id = 600;
			this.mnuVentasReportesPorMarca.Name = "mnuVentasReportesPorMarca";
			this.mnuVentasReportesPorMarca.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesPorMarca_ItemClick);
			// 
			// mnuVentasRVML
			// 
			this.mnuVentasRVML.Caption = "Ventas Por Marca y Local";
			this.mnuVentasRVML.Id = 605;
			this.mnuVentasRVML.Name = "mnuVentasRVML";
			this.mnuVentasRVML.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasRVML_ItemClick);
			// 
			// mnuVentasRepMU
			// 
			this.mnuVentasRepMU.Caption = "Margen de Utilidad";
			this.mnuVentasRepMU.Id = 618;
			this.mnuVentasRepMU.Name = "mnuVentasRepMU";
			this.mnuVentasRepMU.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasRepMU_ItemClick);
			// 
			// mnuVentaRMU
			// 
			this.mnuVentaRMU.Caption = "Resumen Margen de Utilidad";
			this.mnuVentaRMU.Id = 682;
			this.mnuVentaRMU.Name = "mnuVentaRMU";
			this.mnuVentaRMU.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentaRMU_ItemClick);
			// 
			// mnuVentaRMUA
			// 
			this.mnuVentaRMUA.Caption = "Margen Utilidad Por Articulo";
			this.mnuVentaRMUA.Id = 873;
			this.mnuVentaRMUA.Name = "mnuVentaRMUA";
			this.mnuVentaRMUA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentaRMUA_ItemClick);
			// 
			// mnuVentasReportesAIM
			// 
			this.mnuVentasReportesAIM.Caption = "Analisis Información Mensual";
			this.mnuVentasReportesAIM.Id = 648;
			this.mnuVentasReportesAIM.Name = "mnuVentasReportesAIM";
			this.mnuVentasReportesAIM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesAIM_ItemClick);
			// 
			// mnuVentasReportesRF
			// 
			this.mnuVentasReportesRF.Caption = "Revisión de Facturas";
			this.mnuVentasReportesRF.Id = 649;
			this.mnuVentasReportesRF.Name = "mnuVentasReportesRF";
			this.mnuVentasReportesRF.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesRF_ItemClick);
			// 
			// mnuVentasReportesVPD
			// 
			this.mnuVentasReportesVPD.Caption = "Ventas Por Día";
			this.mnuVentasReportesVPD.Id = 825;
			this.mnuVentasReportesVPD.Name = "mnuVentasReportesVPD";
			this.mnuVentasReportesVPD.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesVPD_ItemClick);
			// 
			// mnuVentasReportesFFP
			// 
			this.mnuVentasReportesFFP.Caption = "Numero De Facturas Por Formas De Pago";
			this.mnuVentasReportesFFP.Id = 821;
			this.mnuVentasReportesFFP.Name = "mnuVentasReportesFFP";
			this.mnuVentasReportesFFP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesFFP_ItemClick);
			// 
			// mnuVentasRepGE
			// 
			this.mnuVentasRepGE.Caption = "Garantia Extendida";
			this.mnuVentasRepGE.Id = 714;
			this.mnuVentasRepGE.Name = "mnuVentasRepGE";
			this.mnuVentasRepGE.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasRepGE_ItemClick);
			// 
			// mnuVentasRepRVF
			// 
			this.mnuVentasRepRVF.Caption = "Vendedores Por Fecha";
			this.mnuVentasRepRVF.Id = 764;
			this.mnuVentasRepRVF.Name = "mnuVentasRepRVF";
			this.mnuVentasRepRVF.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasRepRVF_ItemClick);
			// 
			// mnuVentasReportesComparativos
			// 
			this.mnuVentasReportesComparativos.Caption = "Comparativos de Ventas";
			this.mnuVentasReportesComparativos.Id = 732;
			this.mnuVentasReportesComparativos.Name = "mnuVentasReportesComparativos";
			this.mnuVentasReportesComparativos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesComparativos_ItemClick);
			// 
			// mnuVentasReportesVentasPorMes
			// 
			this.mnuVentasReportesVentasPorMes.Caption = "Ventas Anuales Por Local";
			this.mnuVentasReportesVentasPorMes.Id = 743;
			this.mnuVentasReportesVentasPorMes.Name = "mnuVentasReportesVentasPorMes";
			this.mnuVentasReportesVentasPorMes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesVentasPorMes_ItemClick);
			// 
			// mnuReporteVentasMensualesPorLocal
			// 
			this.mnuReporteVentasMensualesPorLocal.Caption = "Ventas Mensuales Por Local";
			this.mnuReporteVentasMensualesPorLocal.Id = 797;
			this.mnuReporteVentasMensualesPorLocal.Name = "mnuReporteVentasMensualesPorLocal";
			this.mnuReporteVentasMensualesPorLocal.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuReporteVentasMensualesPorLocal_ItemClick);
			// 
			// mnuVentasReportesTC
			// 
			this.mnuVentasReportesTC.Caption = "Ventas Tarjetas de Crédito";
			this.mnuVentasReportesTC.Id = 748;
			this.mnuVentasReportesTC.Name = "mnuVentasReportesTC";
			this.mnuVentasReportesTC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesTC_ItemClick);
			// 
			// mnuVentasReportesDTC
			// 
			this.mnuVentasReportesDTC.Caption = "Detalle Tarjetas de Credito";
			this.mnuVentasReportesDTC.Id = 779;
			this.mnuVentasReportesDTC.Name = "mnuVentasReportesDTC";
			this.mnuVentasReportesDTC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesDTC_ItemClick);
			// 
			// mnuVentasReportesEncuestas
			// 
			this.mnuVentasReportesEncuestas.Caption = "Encuestas";
			this.mnuVentasReportesEncuestas.Id = 780;
			this.mnuVentasReportesEncuestas.Name = "mnuVentasReportesEncuestas";
			this.mnuVentasReportesEncuestas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesEncuestas_ItemClick);
			// 
			// mnuVentasReportesCXC
			// 
			this.mnuVentasReportesCXC.Caption = "Resumen de Cuentas Por Cobrar";
			this.mnuVentasReportesCXC.Id = 834;
			this.mnuVentasReportesCXC.Name = "mnuVentasReportesCXC";
			this.mnuVentasReportesCXC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesCXC_ItemClick);
			// 
			// mnuContResCXC
			// 
			this.mnuContResCXC.Caption = "Resumen de Cuentas Por Cobrar";
			this.mnuContResCXC.Id = 546;
			this.mnuContResCXC.Name = "mnuContResCXC";
			this.mnuContResCXC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContResCXC_ItemClick);
			// 
			// mnuVentasReportesVPA
			// 
			this.mnuVentasReportesVPA.Caption = "Vendedor Por Año";
			this.mnuVentasReportesVPA.Id = 839;
			this.mnuVentasReportesVPA.Name = "mnuVentasReportesVPA";
			this.mnuVentasReportesVPA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesVPA_ItemClick);
			// 
			// mnuVentasComisiones
			// 
			this.mnuVentasComisiones.Caption = "Comisiones";
			this.mnuVentasComisiones.Id = 495;
			this.mnuVentasComisiones.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																		 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFacComConfiguracion),
																																																		 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactComCalificacion, true),
																																																		 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactDetalle),
																																																		 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactComConsulta, true)});
			this.mnuVentasComisiones.Name = "mnuVentasComisiones";
			this.mnuVentasComisiones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// MnuFacComConfiguracion
			// 
			this.MnuFacComConfiguracion.Caption = "Configuración";
			this.MnuFacComConfiguracion.Id = 496;
			this.MnuFacComConfiguracion.Name = "MnuFacComConfiguracion";
			this.MnuFacComConfiguracion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFacComConfiguracion_ItemClick);
			// 
			// MnuFactComCalificacion
			// 
			this.MnuFactComCalificacion.Caption = "Calificación";
			this.MnuFactComCalificacion.Id = 497;
			this.MnuFactComCalificacion.Name = "MnuFactComCalificacion";
			this.MnuFactComCalificacion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactComCalificacion_ItemClick);
			// 
			// MnuFactDetalle
			// 
			this.MnuFactDetalle.Caption = "Detalle general";
			this.MnuFactDetalle.Id = 498;
			this.MnuFactDetalle.Name = "MnuFactDetalle";
			this.MnuFactDetalle.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactDetalle_ItemClick);
			// 
			// MnuFactComConsulta
			// 
			this.MnuFactComConsulta.Caption = "Consulta";
			this.MnuFactComConsulta.Id = 499;
			this.MnuFactComConsulta.Name = "MnuFactComConsulta";
			this.MnuFactComConsulta.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactComConsulta_ItemClick);
			// 
			// mnuVentasReportesVentasPorArticulos
			// 
			this.mnuVentasReportesVentasPorArticulos.Caption = "Ventas Por Articulos";
			this.mnuVentasReportesVentasPorArticulos.Id = 846;
			this.mnuVentasReportesVentasPorArticulos.Name = "mnuVentasReportesVentasPorArticulos";
			this.mnuVentasReportesVentasPorArticulos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesVentasPorArticulos_ItemClick);
			// 
			// mnuVentasReportesVentasarticuloPorMes
			// 
			this.mnuVentasReportesVentasarticuloPorMes.Caption = "Ventas De Articulos Por Mes";
			this.mnuVentasReportesVentasarticuloPorMes.Id = 847;
			this.mnuVentasReportesVentasarticuloPorMes.Name = "mnuVentasReportesVentasarticuloPorMes";
			this.mnuVentasReportesVentasarticuloPorMes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasReportesVentasarticuloPorMes_ItemClick);
			// 
			// mnuVentaPromocionesActivas
			// 
			this.mnuVentaPromocionesActivas.Caption = "Promociones Activas";
			this.mnuVentaPromocionesActivas.Id = 603;
			this.mnuVentaPromocionesActivas.Name = "mnuVentaPromocionesActivas";
			this.mnuVentaPromocionesActivas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentaPromocionesActivas_ItemClick);
			// 
			// MnuCredito
			// 
			this.MnuCredito.Caption = "Cré&diPOINT";
			this.MnuCredito.Id = 470;
			this.MnuCredito.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuCreCobranza),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuCreditoCotizador, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPEDC, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPHRC, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPRRC),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentasGCCC, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPRevision, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuCreditoSolRapida),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCrediPointSolicitud),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCrediPointMEF, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPMFCP),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCrediPointMantenedor, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCrediPointVerTerreno),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCrediPointGestiones, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCrediPointCatalogos),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuCreditoCuota, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.MnuCreGastoCobran, false),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPReportes),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzas, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPComisiones),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCartera),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPRD, true),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPVEP),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCrediPOINTRMR),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCRD),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPANDC)});
			this.MnuCredito.Name = "MnuCredito";
			// 
			// MnuCreCobranza
			// 
			this.MnuCreCobranza.Caption = "Cobranza";
			this.MnuCreCobranza.Id = 491;
			this.MnuCreCobranza.ItemShortcut = new DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F11);
			this.MnuCreCobranza.Name = "MnuCreCobranza";
			this.MnuCreCobranza.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCreCobranza_ItemClick);
			// 
			// MnuCreditoCotizador
			// 
			this.MnuCreditoCotizador.Caption = "Cotizador";
			this.MnuCreditoCotizador.Id = 486;
			this.MnuCreditoCotizador.Name = "MnuCreditoCotizador";
			this.MnuCreditoCotizador.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCreditoCotizador_ItemClick);
			// 
			// mnuCPEDC
			// 
			this.mnuCPEDC.Caption = "Envio de Cedulas";
			this.mnuCPEDC.Id = 675;
			this.mnuCPEDC.Name = "mnuCPEDC";
			this.mnuCPEDC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPEDC_ItemClick);
			// 
			// mnuCPHRC
			// 
			this.mnuCPHRC.Caption = "Historial Revisión de Cédulas";
			this.mnuCPHRC.Id = 681;
			this.mnuCPHRC.Name = "mnuCPHRC";
			this.mnuCPHRC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPHRC_ItemClick);
			// 
			// mnuCPRRC
			// 
			this.mnuCPRRC.Caption = "Resumen de Revisión de Cédulas";
			this.mnuCPRRC.Id = 683;
			this.mnuCPRRC.Name = "mnuCPRRC";
			this.mnuCPRRC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPRRC_ItemClick);
			// 
			// mnuVentasGCCC
			// 
			this.mnuVentasGCCC.Caption = "Gestion Cartera de Clientes";
			this.mnuVentasGCCC.Id = 786;
			this.mnuVentasGCCC.Name = "mnuVentasGCCC";
			this.mnuVentasGCCC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasGCCC_ItemClick);
			// 
			// mnuCPRevision
			// 
			this.mnuCPRevision.Caption = "Revisión De Cedulas";
			this.mnuCPRevision.Id = 883;
			this.mnuCPRevision.Name = "mnuCPRevision";
			this.mnuCPRevision.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPRevision_ItemClick);
			// 
			// MnuCreditoSolRapida
			// 
			this.MnuCreditoSolRapida.Caption = "Solicitud Rápida";
			this.MnuCreditoSolRapida.Id = 471;
			this.MnuCreditoSolRapida.Name = "MnuCreditoSolRapida";
			this.MnuCreditoSolRapida.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.MnuCreditoSolRapida.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCreditoSolRapida_ItemClick);
			// 
			// mnuCrediPointSolicitud
			// 
			this.mnuCrediPointSolicitud.Caption = "Solicitud";
			this.mnuCrediPointSolicitud.Id = 474;
			this.mnuCrediPointSolicitud.Name = "mnuCrediPointSolicitud";
			this.mnuCrediPointSolicitud.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuCrediPointSolicitud.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCrediPointSolicitud_ItemClick);
			// 
			// mnuCrediPointMEF
			// 
			this.mnuCrediPointMEF.Caption = "Modifica Entidad Financiera";
			this.mnuCrediPointMEF.Id = 607;
			this.mnuCrediPointMEF.Name = "mnuCrediPointMEF";
			this.mnuCrediPointMEF.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCrediPointMEF_ItemClick);
			// 
			// mnuCPMFCP
			// 
			this.mnuCPMFCP.Caption = "Modifica Fecha De Pago";
			this.mnuCPMFCP.Id = 617;
			this.mnuCPMFCP.Name = "mnuCPMFCP";
			this.mnuCPMFCP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPMFCP_ItemClick);
			// 
			// mnuCrediPointMantenedor
			// 
			this.mnuCrediPointMantenedor.Caption = "Mantenedor";
			this.mnuCrediPointMantenedor.Id = 475;
			this.mnuCrediPointMantenedor.Name = "mnuCrediPointMantenedor";
			this.mnuCrediPointMantenedor.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuCrediPointVerTerreno
			// 
			this.mnuCrediPointVerTerreno.Caption = "Verificación Terreno";
			this.mnuCrediPointVerTerreno.Id = 476;
			this.mnuCrediPointVerTerreno.Name = "mnuCrediPointVerTerreno";
			this.mnuCrediPointVerTerreno.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuCrediPointGestiones
			// 
			this.mnuCrediPointGestiones.Caption = "Gestiones";
			this.mnuCrediPointGestiones.Id = 477;
			this.mnuCrediPointGestiones.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																				new DevExpress.XtraBars.LinkPersistInfo(this.mnuCrediPointGestVerCed),
																																																				new DevExpress.XtraBars.LinkPersistInfo(this.mnuCrediPointGestRevCed),
																																																				new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem51)});
			this.mnuCrediPointGestiones.Name = "mnuCrediPointGestiones";
			this.mnuCrediPointGestiones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuCrediPointGestVerCed
			// 
			this.mnuCrediPointGestVerCed.Caption = "Verificar CEDULA";
			this.mnuCrediPointGestVerCed.Id = 478;
			this.mnuCrediPointGestVerCed.Name = "mnuCrediPointGestVerCed";
			// 
			// mnuCrediPointGestRevCed
			// 
			this.mnuCrediPointGestRevCed.Caption = "Revisa CEDULA Supervisor";
			this.mnuCrediPointGestRevCed.Id = 481;
			this.mnuCrediPointGestRevCed.Name = "mnuCrediPointGestRevCed";
			// 
			// barButtonItem51
			// 
			this.barButtonItem51.Caption = "barButtonItem51";
			this.barButtonItem51.Id = 795;
			this.barButtonItem51.Name = "barButtonItem51";
			// 
			// mnuCrediPointCatalogos
			// 
			this.mnuCrediPointCatalogos.Caption = "Catálogos";
			this.mnuCrediPointCatalogos.Id = 480;
			this.mnuCrediPointCatalogos.Name = "mnuCrediPointCatalogos";
			this.mnuCrediPointCatalogos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// MnuCreditoCuota
			// 
			this.MnuCreditoCuota.Caption = "Parámetros de Crédito";
			this.MnuCreditoCuota.Id = 489;
			this.MnuCreditoCuota.Name = "MnuCreditoCuota";
			this.MnuCreditoCuota.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCreditoCuota_ItemClick);
			// 
			// MnuCreGastoCobran
			// 
			this.MnuCreGastoCobran.Caption = "Gestión de Cobranza";
			this.MnuCreGastoCobran.Id = 490;
			this.MnuCreGastoCobran.Name = "MnuCreGastoCobran";
			this.MnuCreGastoCobran.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.MnuCreGastoCobran.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCreGastoCobran_ItemClick);
			// 
			// mnuCPReportes
			// 
			this.mnuCPReportes.Caption = "Reportes";
			this.mnuCPReportes.Id = 596;
			this.mnuCPReportes.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPReportesCliPoint),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPLCS),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPReportesEC),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPRepEquifax),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPRepCP),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPRepRDD),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPRepECP),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem37),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPReportesRCV),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPReportesGDCM),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPReportesAbonos),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPReportesNDC),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPReportesBR),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem76),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPReportesREC),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCrediPOINTReportesRMR),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPReportesRRM),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPReportesCVPM),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCRResumenCredito)});
			this.mnuCPReportes.Name = "mnuCPReportes";
			// 
			// mnuCPReportesCliPoint
			// 
			this.mnuCPReportesCliPoint.Caption = "Clientes POINT";
			this.mnuCPReportesCliPoint.Id = 597;
			this.mnuCPReportesCliPoint.Name = "mnuCPReportesCliPoint";
			this.mnuCPReportesCliPoint.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPReportesCliPoint_ItemClick);
			// 
			// mnuCPLCS
			// 
			this.mnuCPLCS.Caption = "Lista de Operaciones de Credito";
			this.mnuCPLCS.Id = 615;
			this.mnuCPLCS.Name = "mnuCPLCS";
			this.mnuCPLCS.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPLCS_ItemClick);
			// 
			// mnuCPReportesEC
			// 
			this.mnuCPReportesEC.Caption = "Estado de Cartera";
			this.mnuCPReportesEC.Id = 835;
			this.mnuCPReportesEC.Name = "mnuCPReportesEC";
			this.mnuCPReportesEC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPReportesEC_ItemClick);
			// 
			// mnuCPRepEquifax
			// 
			this.mnuCPRepEquifax.Caption = "Equifax";
			this.mnuCPRepEquifax.Id = 616;
			this.mnuCPRepEquifax.Name = "mnuCPRepEquifax";
			this.mnuCPRepEquifax.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPRepEquifax_ItemClick);
			// 
			// mnuCPRepCP
			// 
			this.mnuCPRepCP.Caption = "Pagos Pichincha";
			this.mnuCPRepCP.Id = 620;
			this.mnuCPRepCP.Name = "mnuCPRepCP";
			this.mnuCPRepCP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPRepCP_ItemClick);
			// 
			// mnuCPRepRDD
			// 
			this.mnuCPRepRDD.Caption = "Recepcion de Documentos";
			this.mnuCPRepRDD.Id = 671;
			this.mnuCPRepRDD.Name = "mnuCPRepRDD";
			this.mnuCPRepRDD.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPRepRDD_ItemClick);
			// 
			// mnuCPRepECP
			// 
			this.mnuCPRepECP.Caption = "Estado CrediPOINT";
			this.mnuCPRepECP.Id = 674;
			this.mnuCPRepECP.Name = "mnuCPRepECP";
			this.mnuCPRepECP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPRepECP_ItemClick);
			// 
			// barButtonItem37
			// 
			this.barButtonItem37.Caption = "Retiros de Mercaderia";
			this.barButtonItem37.Id = 684;
			this.barButtonItem37.Name = "barButtonItem37";
			this.barButtonItem37.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem37_ItemClick);
			// 
			// mnuCPReportesRCV
			// 
			this.mnuCPReportesRCV.Caption = "Cartera Vendida";
			this.mnuCPReportesRCV.Id = 702;
			this.mnuCPReportesRCV.Name = "mnuCPReportesRCV";
			this.mnuCPReportesRCV.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPReportesRCV_ItemClick);
			// 
			// mnuCPReportesGDCM
			// 
			this.mnuCPReportesGDCM.Caption = "Gestión de Cobranza Mensual";
			this.mnuCPReportesGDCM.Id = 733;
			this.mnuCPReportesGDCM.Name = "mnuCPReportesGDCM";
			this.mnuCPReportesGDCM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPReportesGDCM_ItemClick);
			// 
			// mnuCPReportesAbonos
			// 
			this.mnuCPReportesAbonos.Caption = "Anticipos - Abonos";
			this.mnuCPReportesAbonos.Id = 742;
			this.mnuCPReportesAbonos.Name = "mnuCPReportesAbonos";
			this.mnuCPReportesAbonos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPReportesAbonos_ItemClick);
			// 
			// mnuCPReportesNDC
			// 
			this.mnuCPReportesNDC.Caption = "Notas de Créditos";
			this.mnuCPReportesNDC.Id = 744;
			this.mnuCPReportesNDC.Name = "mnuCPReportesNDC";
			this.mnuCPReportesNDC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPReportesNDC_ItemClick);
			// 
			// mnuCPReportesBR
			// 
			this.mnuCPReportesBR.Caption = "Balance de Retiros de Mercaderia";
			this.mnuCPReportesBR.Id = 789;
			this.mnuCPReportesBR.Name = "mnuCPReportesBR";
			this.mnuCPReportesBR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPReportesBR_ItemClick);
			// 
			// barButtonItem76
			// 
			this.barButtonItem76.Caption = "Reporte de Retiros";
			this.barButtonItem76.Id = 787;
			this.barButtonItem76.Name = "barButtonItem76";
			this.barButtonItem76.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem76_ItemClick);
			// 
			// mnuCPReportesREC
			// 
			this.mnuCPReportesREC.Caption = "Resumen Estado de Cartera";
			this.mnuCPReportesREC.Id = 810;
			this.mnuCPReportesREC.Name = "mnuCPReportesREC";
			this.mnuCPReportesREC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPReportesREC_ItemClick);
			// 
			// mnuCrediPOINTReportesRMR
			// 
			this.mnuCrediPOINTReportesRMR.Caption = "Recepción Mercaderia de Retiro";
			this.mnuCrediPOINTReportesRMR.Id = 833;
			this.mnuCrediPOINTReportesRMR.Name = "mnuCrediPOINTReportesRMR";
			this.mnuCrediPOINTReportesRMR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCrediPOINTReportesRMR_ItemClick);
			// 
			// mnuCPReportesRRM
			// 
			this.mnuCPReportesRRM.Caption = "Resumen Retiros";
			this.mnuCPReportesRRM.Id = 864;
			this.mnuCPReportesRRM.Name = "mnuCPReportesRRM";
			this.mnuCPReportesRRM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPReportesRRM_ItemClick);
			// 
			// mnuCPReportesCVPM
			// 
			this.mnuCPReportesCVPM.Caption = "Cartera Vencida Por Mes";
			this.mnuCPReportesCVPM.Id = 867;
			this.mnuCPReportesCVPM.Name = "mnuCPReportesCVPM";
			this.mnuCPReportesCVPM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPReportesCVPM_ItemClick);
			// 
			// mnuCRResumenCredito
			// 
			this.mnuCRResumenCredito.Caption = "Resumen Credito";
			this.mnuCRResumenCredito.Id = 972;
			this.mnuCRResumenCredito.Name = "mnuCRResumenCredito";
			this.mnuCRResumenCredito.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCRResumenCredito_ItemClick);
			// 
			// mnuCPCobranzas
			// 
			this.mnuCPCobranzas.Caption = "Cobranzas";
			this.mnuCPCobranzas.Id = 696;
			this.mnuCPCobranzas.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzasGDC),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzasCboOpe, true),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPInformeCobrador),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPIO),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCobranzasRGC, true),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCP),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzasRCP),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem50),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzasCCPPC),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem72, true),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzasLPCA),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzasEDC),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzasMRPR),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzasRPE),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobCM),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzasRPP)});
			this.mnuCPCobranzas.Name = "mnuCPCobranzas";
			// 
			// mnuCPCobranzasGDC
			// 
			this.mnuCPCobranzasGDC.Caption = "Gestor De Cobranzas";
			this.mnuCPCobranzasGDC.Id = 871;
			this.mnuCPCobranzasGDC.Name = "mnuCPCobranzasGDC";
			this.mnuCPCobranzasGDC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobranzasGDC_ItemClick);
			// 
			// mnuCPCobranzasCboOpe
			// 
			this.mnuCPCobranzasCboOpe.Caption = "Informe Cobrador/Operador";
			this.mnuCPCobranzasCboOpe.Id = 981;
			this.mnuCPCobranzasCboOpe.Name = "mnuCPCobranzasCboOpe";
			this.mnuCPCobranzasCboOpe.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobranzasCboOpe_ItemClick);
			// 
			// mnuCPInformeCobrador
			// 
			this.mnuCPInformeCobrador.Caption = "Informes Cobrador";
			this.mnuCPInformeCobrador.Id = 598;
			this.mnuCPInformeCobrador.Name = "mnuCPInformeCobrador";
			this.mnuCPInformeCobrador.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPInformeCobrador_ItemClick);
			// 
			// mnuCPIO
			// 
			this.mnuCPIO.Caption = "Informes Operador";
			this.mnuCPIO.Id = 609;
			this.mnuCPIO.Name = "mnuCPIO";
			this.mnuCPIO.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPIO_ItemClick);
			// 
			// mnuCobranzasRGC
			// 
			this.mnuCobranzasRGC.Caption = "Gestiones de Crédito";
			this.mnuCobranzasRGC.Id = 717;
			this.mnuCobranzasRGC.Name = "mnuCobranzasRGC";
			this.mnuCobranzasRGC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCobranzasRGC_ItemClick);
			// 
			// mnuCPCP
			// 
			this.mnuCPCP.Caption = "Compromisos de Pago";
			this.mnuCPCP.Id = 608;
			this.mnuCPCP.Name = "mnuCPCP";
			this.mnuCPCP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCP_ItemClick);
			// 
			// mnuCPCobranzasRCP
			// 
			this.mnuCPCobranzasRCP.Caption = "Resumen Cartera Pichincha";
			this.mnuCPCobranzasRCP.Id = 697;
			this.mnuCPCobranzasRCP.Name = "mnuCPCobranzasRCP";
			this.mnuCPCobranzasRCP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobranzasRCP_ItemClick);
			// 
			// barButtonItem50
			// 
			this.barButtonItem50.Caption = "Comparación Cobranza Pichincha";
			this.barButtonItem50.Id = 735;
			this.barButtonItem50.Name = "barButtonItem50";
			this.barButtonItem50.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem50_ItemClick);
			// 
			// mnuCPCobranzasCCPPC
			// 
			this.mnuCPCobranzasCCPPC.Caption = "Comparación Cartera Pichincha Por Cuotas";
			this.mnuCPCobranzasCCPPC.Id = 838;
			this.mnuCPCobranzasCCPPC.Name = "mnuCPCobranzasCCPPC";
			this.mnuCPCobranzasCCPPC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobranzasCCPPC_ItemClick);
			// 
			// barButtonItem72
			// 
			this.barButtonItem72.Caption = "Lista de Clientes de Envio Mensajes de Texto";
			this.barButtonItem72.Id = 791;
			this.barButtonItem72.Name = "barButtonItem72";
			this.barButtonItem72.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem72_ItemClick);
			// 
			// mnuCPCobranzasLPCA
			// 
			this.mnuCPCobranzasLPCA.Caption = "Lista De Pagos Clientes Atrasados con Saldo";
			this.mnuCPCobranzasLPCA.Id = 803;
			this.mnuCPCobranzasLPCA.Name = "mnuCPCobranzasLPCA";
			this.mnuCPCobranzasLPCA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobranzasLPCA_ItemClick);
			// 
			// mnuCPCobranzasEDC
			// 
			this.mnuCPCobranzasEDC.Caption = "Estadisticas De Cartera";
			this.mnuCPCobranzasEDC.Id = 850;
			this.mnuCPCobranzasEDC.Name = "mnuCPCobranzasEDC";
			this.mnuCPCobranzasEDC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobranzasEDC_ItemClick);
			// 
			// mnuCPCobranzasMRPR
			// 
			this.mnuCPCobranzasMRPR.Caption = "Mercaderia Reservada Por Retiro";
			this.mnuCPCobranzasMRPR.Id = 869;
			this.mnuCPCobranzasMRPR.Name = "mnuCPCobranzasMRPR";
			this.mnuCPCobranzasMRPR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobranzasMRPR_ItemClick);
			// 
			// mnuCPCobranzasRPE
			// 
			this.mnuCPCobranzasRPE.Caption = "Refinanciaciones Por Estado";
			this.mnuCPCobranzasRPE.Id = 886;
			this.mnuCPCobranzasRPE.Name = "mnuCPCobranzasRPE";
			this.mnuCPCobranzasRPE.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobranzasRPE_ItemClick);
			// 
			// mnuCPCobCM
			// 
			this.mnuCPCobCM.Caption = "Cash Management";
			this.mnuCPCobCM.Id = 970;
			this.mnuCPCobCM.Name = "mnuCPCobCM";
			this.mnuCPCobCM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobCM_ItemClick);
			// 
			// mnuCPCobranzasRPP
			// 
			this.mnuCPCobranzasRPP.Caption = "Retorna Pichincha Point";
			this.mnuCPCobranzasRPP.Id = 1000;
			this.mnuCPCobranzasRPP.Name = "mnuCPCobranzasRPP";
			this.mnuCPCobranzasRPP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobranzasRPP_ItemClick);
			// 
			// mnuCPComisiones
			// 
			this.mnuCPComisiones.Caption = "Comisiones";
			this.mnuCPComisiones.Id = 978;
			this.mnuCPComisiones.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPComisionesAsigCobr),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPComisionesTrmOpe),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPComisionesCom),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCreditoGesCall),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.mnuCreditoProReal),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem84)});
			this.mnuCPComisiones.Name = "mnuCPComisiones";
			// 
			// mnuCPComisionesAsigCobr
			// 
			this.mnuCPComisionesAsigCobr.Caption = "Asignación Cobradores";
			this.mnuCPComisionesAsigCobr.Id = 979;
			this.mnuCPComisionesAsigCobr.Name = "mnuCPComisionesAsigCobr";
			this.mnuCPComisionesAsigCobr.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPComisionesAsigCobr_ItemClick);
			// 
			// mnuCPComisionesTrmOpe
			// 
			this.mnuCPComisionesTrmOpe.Caption = "Tramas Operadoras";
			this.mnuCPComisionesTrmOpe.Id = 980;
			this.mnuCPComisionesTrmOpe.Name = "mnuCPComisionesTrmOpe";
			this.mnuCPComisionesTrmOpe.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPComisionesTrmOpe_ItemClick);
			// 
			// mnuCPComisionesCom
			// 
			this.mnuCPComisionesCom.Caption = "Porcentaje Comisiones";
			this.mnuCPComisionesCom.Id = 982;
			this.mnuCPComisionesCom.Name = "mnuCPComisionesCom";
			this.mnuCPComisionesCom.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPComisionesCom_ItemClick);
			// 
			// mnuCreditoGesCall
			// 
			this.mnuCreditoGesCall.Caption = "Gestiones Call Center";
			this.mnuCreditoGesCall.Id = 983;
			this.mnuCreditoGesCall.Name = "mnuCreditoGesCall";
			this.mnuCreditoGesCall.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCreditoGesCall_ItemClick);
			// 
			// mnuCreditoProReal
			// 
			this.mnuCreditoProReal.Caption = "Reporte Cobranza Proyección";
			this.mnuCreditoProReal.Id = 984;
			this.mnuCreditoProReal.Name = "mnuCreditoProReal";
			this.mnuCreditoProReal.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCreditoProReal_ItemClick);
			// 
			// barButtonItem84
			// 
			this.barButtonItem84.Caption = "Resumen de Comisiones";
			this.barButtonItem84.Id = 988;
			this.barButtonItem84.Name = "barButtonItem84";
			this.barButtonItem84.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem84_ItemClick);
			// 
			// mnuCPCartera
			// 
			this.mnuCPCartera.Caption = "Cartera";
			this.mnuCPCartera.Id = 724;
			this.mnuCPCartera.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																															new DevExpress.XtraBars.LinkPersistInfo(this.barSubItem10),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.barSubItem11),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzasCarteraAIP),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCobranzasCERTIFICADO)});
			this.mnuCPCartera.Name = "mnuCPCartera";
			// 
			// barSubItem10
			// 
			this.barSubItem10.Caption = "Solidario";
			this.barSubItem10.Id = 727;
			this.barSubItem10.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuCrediPointDFG),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuCreLotesUnibanco),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuContRepAnt),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPRRCS),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuContCCxCCP),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem57)});
			this.barSubItem10.Name = "barSubItem10";
			// 
			// mnuCrediPointDFG
			// 
			this.mnuCrediPointDFG.Caption = "Debitos Fondo De Garantia";
			this.mnuCrediPointDFG.Id = 606;
			this.mnuCrediPointDFG.Name = "mnuCrediPointDFG";
			this.mnuCrediPointDFG.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCrediPointDFG_ItemClick);
			// 
			// mnuCreLotesUnibanco
			// 
			this.mnuCreLotesUnibanco.Caption = "Lotes Solidario";
			this.mnuCreLotesUnibanco.Id = 515;
			this.mnuCreLotesUnibanco.Name = "mnuCreLotesUnibanco";
			this.mnuCreLotesUnibanco.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCreLotesUnibanco_ItemClick);
			// 
			// mnuContRepAnt
			// 
			this.mnuContRepAnt.Caption = "Reposición Cartera Solidario";
			this.mnuContRepAnt.Id = 548;
			this.mnuContRepAnt.Name = "mnuContRepAnt";
			this.mnuContRepAnt.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContRepAnt_ItemClick);
			// 
			// mnuCPRRCS
			// 
			this.mnuCPRRCS.Caption = "Reporte Reposicion Cartera Solidario";
			this.mnuCPRRCS.Id = 667;
			this.mnuCPRRCS.Name = "mnuCPRRCS";
			this.mnuCPRRCS.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPRRCS_ItemClick);
			// 
			// mnuContCCxCCP
			// 
			this.mnuContCCxCCP.Caption = "Liquidación Cartera Solidaio";
			this.mnuContCCxCCP.Id = 553;
			this.mnuContCCxCCP.Name = "mnuContCCxCCP";
			this.mnuContCCxCCP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuContCCxCCP_ItemClick);
			// 
			// barButtonItem57
			// 
			this.barButtonItem57.Caption = "barButtonItem57";
			this.barButtonItem57.Id = 750;
			this.barButtonItem57.Name = "barButtonItem57";
			this.barButtonItem57.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.barButtonItem57.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem57_ItemClick);
			// 
			// barSubItem11
			// 
			this.barSubItem11.Caption = "Pichincha";
			this.barSubItem11.Id = 729;
			this.barSubItem11.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuCpLbp),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPLotesPPagos),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPDFGP),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem82)});
			this.barSubItem11.Name = "barSubItem11";
			// 
			// mnuCpLbp
			// 
			this.mnuCpLbp.Caption = "Liquidación";
			this.mnuCpLbp.Id = 647;
			this.mnuCpLbp.Name = "mnuCpLbp";
			this.mnuCpLbp.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCpLbp_ItemClick);
			// 
			// mnuCPLotesPPagos
			// 
			this.mnuCPLotesPPagos.Caption = "Pagos";
			this.mnuCPLotesPPagos.Id = 730;
			this.mnuCPLotesPPagos.Name = "mnuCPLotesPPagos";
			this.mnuCPLotesPPagos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPLotesPPagos_ItemClick);
			// 
			// mnuCPDFGP
			// 
			this.mnuCPDFGP.Caption = "Debitos Fondo de Garantia";
			this.mnuCPDFGP.Id = 705;
			this.mnuCPDFGP.Name = "mnuCPDFGP";
			this.mnuCPDFGP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPDFGP_ItemClick);
			// 
			// barButtonItem82
			// 
			this.barButtonItem82.Caption = "Generacion";
			this.barButtonItem82.Id = 973;
			this.barButtonItem82.Name = "barButtonItem82";
			this.barButtonItem82.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem82_ItemClick);
			// 
			// mnuCPCobranzasCarteraAIP
			// 
			this.mnuCPCobranzasCarteraAIP.Caption = "Asignacion Cartera I P";
			this.mnuCPCobranzasCarteraAIP.Id = 885;
			this.mnuCPCobranzasCarteraAIP.Name = "mnuCPCobranzasCarteraAIP";
			this.mnuCPCobranzasCarteraAIP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobranzasCarteraAIP_ItemClick);
			// 
			// mnuCPCobranzasCERTIFICADO
			// 
			this.mnuCPCobranzasCERTIFICADO.Caption = "Certificado";
			this.mnuCPCobranzasCERTIFICADO.Id = 986;
			this.mnuCPCobranzasCERTIFICADO.Name = "mnuCPCobranzasCERTIFICADO";
			this.mnuCPCobranzasCERTIFICADO.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCobranzasCERTIFICADO_ItemClick);
			// 
			// mnuCPRD
			// 
			this.mnuCPRD.Caption = "Recepcion de Documentos";
			this.mnuCPRD.Id = 670;
			this.mnuCPRD.Name = "mnuCPRD";
			this.mnuCPRD.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPRD_ItemClick);
			// 
			// mnuCPVEP
			// 
			this.mnuCPVEP.Caption = "Verificación Entrega de Productos";
			this.mnuCPVEP.Id = 672;
			this.mnuCPVEP.Name = "mnuCPVEP";
			this.mnuCPVEP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPVEP_ItemClick);
			// 
			// mnuCrediPOINTRMR
			// 
			this.mnuCrediPOINTRMR.Caption = "Recepción Mercaderia De Retiro";
			this.mnuCrediPOINTRMR.Id = 832;
			this.mnuCrediPOINTRMR.Name = "mnuCrediPOINTRMR";
			this.mnuCrediPOINTRMR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCrediPOINTRMR_ItemClick);
			// 
			// mnuCPCRD
			// 
			this.mnuCPCRD.Caption = "Credito Directo";
			this.mnuCPCRD.Id = 889;
			this.mnuCPCRD.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.mnuCPCDRDC, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.Caption),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCDRVC, true),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCDSR),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCDRSLT),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuCPCDLCPG, true)});
			this.mnuCPCRD.Name = "mnuCPCRD";
			// 
			// mnuCPCDRDC
			// 
			this.mnuCPCDRDC.Caption = "Recepción de Cedulas";
			this.mnuCPCDRDC.Id = 890;
			this.mnuCPCDRDC.Name = "mnuCPCDRDC";
			this.mnuCPCDRDC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCDRDC_ItemClick);
			// 
			// mnuCPCDRVC
			// 
			this.mnuCPCDRVC.Caption = "Revisión de Cedulas";
			this.mnuCPCDRVC.Id = 891;
			this.mnuCPCDRVC.Name = "mnuCPCDRVC";
			this.mnuCPCDRVC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCDRVC_ItemClick);
			// 
			// mnuCPCDSR
			// 
			this.mnuCPCDSR.Caption = "Solicitud Rapida";
			this.mnuCPCDSR.Id = 892;
			this.mnuCPCDSR.Name = "mnuCPCDSR";
			this.mnuCPCDSR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCDSR_ItemClick);
			// 
			// mnuCPCDRSLT
			// 
			this.mnuCPCDRSLT.Caption = "Solicitudes";
			this.mnuCPCDRSLT.Id = 893;
			this.mnuCPCDRSLT.Name = "mnuCPCDRSLT";
			this.mnuCPCDRSLT.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCDRSLT_ItemClick);
			// 
			// mnuCPCDLCPG
			// 
			this.mnuCPCDLCPG.Caption = "Lista Clientes Para Gestión";
			this.mnuCPCDLCPG.Id = 894;
			this.mnuCPCDLCPG.Name = "mnuCPCDLCPG";
			this.mnuCPCDLCPG.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPCDLCPG_ItemClick);
			// 
			// mnuCPANDC
			// 
			this.mnuCPANDC.Caption = "Aprobación Notas De Crédito";
			this.mnuCPANDC.Id = 989;
			this.mnuCPANDC.Name = "mnuCPANDC";
			this.mnuCPANDC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuCPANDC_ItemClick);
			// 
			// MnuProduc
			// 
			this.MnuProduc.Caption = "&Producción";
			this.MnuProduc.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.MnuProduc.Id = 147;
			this.MnuProduc.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuProducRecetas),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuProducOrdProd),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.MnuProdEstado, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuProduccionPendientes),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem18),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuProduccionTPT),
																																													 new DevExpress.XtraBars.LinkPersistInfo(this.mnuGPPD),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.MnuProducProdTerm, false),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuProdLiquida, false),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.MnuProducConsReal, false),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.MnuProducActualiza, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.MnuProducReportes, true),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.MnuProdClientes, false),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.MnuProdSemana, false),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.btProdCosto, false),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuProdGeneraC, false),
																																													 new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuProdAgro, false)});
			this.MnuProduc.Name = "MnuProduc";
			// 
			// MnuProducRecetas
			// 
			this.MnuProducRecetas.Caption = "Recetas";
			this.MnuProducRecetas.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.MnuProducRecetas.Id = 148;
			this.MnuProducRecetas.ImageIndex = 35;
			this.MnuProducRecetas.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F2));
			this.MnuProducRecetas.Name = "MnuProducRecetas";
			this.MnuProducRecetas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuProducRecetas_ItemClick);
			// 
			// MnuProducOrdProd
			// 
			this.MnuProducOrdProd.Caption = "Orden de Producción";
			this.MnuProducOrdProd.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.MnuProducOrdProd.Id = 149;
			this.MnuProducOrdProd.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F3));
			this.MnuProducOrdProd.Name = "MnuProducOrdProd";
			this.MnuProducOrdProd.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuProducOrdProd_ItemClick);
			// 
			// MnuProdEstado
			// 
			this.MnuProdEstado.Caption = "Estado Producción";
			this.MnuProdEstado.Id = 513;
			this.MnuProdEstado.Name = "MnuProdEstado";
			this.MnuProdEstado.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuProdEstado_ItemClick);
			// 
			// mnuProduccionPendientes
			// 
			this.mnuProduccionPendientes.Caption = "Pendientes Produccion";
			this.mnuProduccionPendientes.Id = 514;
			this.mnuProduccionPendientes.Name = "mnuProduccionPendientes";
			this.mnuProduccionPendientes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuProduccionPendientes_ItemClick);
			// 
			// barButtonItem18
			// 
			this.barButtonItem18.Caption = "Pendientes de Produccion Portatiles";
			this.barButtonItem18.Id = 593;
			this.barButtonItem18.Name = "barButtonItem18";
			this.barButtonItem18.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem18_ItemClick);
			// 
			// mnuProduccionTPT
			// 
			this.mnuProduccionTPT.Caption = "Transferencia Producción Terminada";
			this.mnuProduccionTPT.Id = 580;
			this.mnuProduccionTPT.Name = "mnuProduccionTPT";
			this.mnuProduccionTPT.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuProduccionTPT_ItemClick);
			// 
			// mnuGPPD
			// 
			this.mnuGPPD.Caption = "Genera Transferencia de Partes Dañadas";
			this.mnuGPPD.Id = 594;
			this.mnuGPPD.Name = "mnuGPPD";
			this.mnuGPPD.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuGPPD_ItemClick);
			// 
			// MnuProducProdTerm
			// 
			this.MnuProducProdTerm.Caption = "Producto Terminado";
			this.MnuProducProdTerm.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.MnuProducProdTerm.Id = 150;
			this.MnuProducProdTerm.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F4));
			this.MnuProducProdTerm.Name = "MnuProducProdTerm";
			this.MnuProducProdTerm.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuProducProdTerm_ItemClick);
			// 
			// mnuProdLiquida
			// 
			this.mnuProdLiquida.Caption = "Liquidación Órden Producción";
			this.mnuProdLiquida.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.mnuProdLiquida.Id = 292;
			this.mnuProdLiquida.Name = "mnuProdLiquida";
			this.mnuProdLiquida.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuProdLiquida_ItemClick);
			// 
			// MnuProducConsReal
			// 
			this.MnuProducConsReal.Caption = "Consumo Real";
			this.MnuProducConsReal.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.MnuProducConsReal.Id = 151;
			this.MnuProducConsReal.ImageIndex = 36;
			this.MnuProducConsReal.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F5));
			this.MnuProducConsReal.Name = "MnuProducConsReal";
			this.MnuProducConsReal.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuProducConsReal_ItemClick);
			// 
			// MnuProducActualiza
			// 
			this.MnuProducActualiza.Caption = "Actualización de Precios";
			this.MnuProducActualiza.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.MnuProducActualiza.Id = 152;
			this.MnuProducActualiza.Name = "MnuProducActualiza";
			this.MnuProducActualiza.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuProducActualiza_ItemClick);
			// 
			// MnuProducReportes
			// 
			this.MnuProducReportes.Caption = "Reportes";
			this.MnuProducReportes.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.MnuProducReportes.Id = 154;
			this.MnuProducReportes.Name = "MnuProducReportes";
			this.MnuProducReportes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuProducReportes_ItemClick);
			// 
			// MnuProdClientes
			// 
			this.MnuProdClientes.Caption = "Presupuesto de Clientes";
			this.MnuProdClientes.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.MnuProdClientes.Id = 203;
			this.MnuProdClientes.Name = "MnuProdClientes";
			this.MnuProdClientes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuProdClientes_ItemClick);
			// 
			// MnuProdSemana
			// 
			this.MnuProdSemana.Caption = "Producción Semanal";
			this.MnuProdSemana.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.MnuProdSemana.Id = 209;
			this.MnuProdSemana.Name = "MnuProdSemana";
			this.MnuProdSemana.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem9_ItemClick);
			// 
			// btProdCosto
			// 
			this.btProdCosto.Caption = "Costos de Producción";
			this.btProdCosto.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.btProdCosto.Id = 299;
			this.btProdCosto.Name = "btProdCosto";
			this.btProdCosto.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btProdCosto_ItemClick);
			// 
			// mnuProdGeneraC
			// 
			this.mnuProdGeneraC.Caption = "Generación de Consumo";
			this.mnuProdGeneraC.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.mnuProdGeneraC.Id = 303;
			this.mnuProdGeneraC.Name = "mnuProdGeneraC";
			this.mnuProdGeneraC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuProdGeneraC_ItemClick);
			// 
			// mnuProdAgro
			// 
			this.mnuProdAgro.Caption = "Actividades Agrícolas";
			this.mnuProdAgro.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.mnuProdAgro.Id = 307;
			this.mnuProdAgro.Name = "mnuProdAgro";
			this.mnuProdAgro.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuProdAgro_ItemClick);
			// 
			// mnuDT
			// 
			this.mnuDT.Caption = "Departamento &Técnico";
			this.mnuDT.Id = 520;
			this.mnuDT.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.MnuDTSolicitud, true),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIREP),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTRA),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIREPC),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIngresoMultiple, true),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTRetiro, true),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRM),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRMP),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTProductosDeFeria, true),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRAT, true),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRSNC),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTAprobacionRemates),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRRPT, true),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRRMA),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.frmDTTransfLocales),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRAPD, true),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTASSF),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRREP, true),
																																											 new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRCDE)});
			this.mnuDT.Name = "mnuDT";
			// 
			// MnuDTSolicitud
			// 
			this.MnuDTSolicitud.Caption = "Ingreso Departamento Técnico";
			this.MnuDTSolicitud.Id = 522;
			this.MnuDTSolicitud.Name = "MnuDTSolicitud";
			this.MnuDTSolicitud.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuDTSolicitud_ItemClick);
			// 
			// mnuDTIREP
			// 
			this.mnuDTIREP.Caption = "Transferencia Servicio Tecnico";
			this.mnuDTIREP.Id = 737;
			this.mnuDTIREP.Name = "mnuDTIREP";
			this.mnuDTIREP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem52_ItemClick);
			// 
			// mnuDTRA
			// 
			this.mnuDTRA.Caption = "Recepción de Articulos";
			this.mnuDTRA.Id = 801;
			this.mnuDTRA.Name = "mnuDTRA";
			this.mnuDTRA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuDTRA_ItemClick);
			// 
			// mnuDTIREPC
			// 
			this.mnuDTIREPC.Caption = "Entrega de Articulos";
			this.mnuDTIREPC.Id = 739;
			this.mnuDTIREPC.Name = "mnuDTIREPC";
			this.mnuDTIREPC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem54_ItemClick);
			// 
			// mnuDTIngresoMultiple
			// 
			this.mnuDTIngresoMultiple.Caption = "Ingreso Multiple";
			this.mnuDTIngresoMultiple.Id = 562;
			this.mnuDTIngresoMultiple.Name = "mnuDTIngresoMultiple";
			this.mnuDTIngresoMultiple.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuDTIngresoMultiple_ItemClick);
			// 
			// mnuDTRetiro
			// 
			this.mnuDTRetiro.Caption = "Ingreso Por Retiro de Mercaderia";
			this.mnuDTRetiro.Id = 583;
			this.mnuDTRetiro.Name = "mnuDTRetiro";
			this.mnuDTRetiro.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuDTRetiro_ItemClick);
			// 
			// mnuDTIRM
			// 
			this.mnuDTIRM.Caption = "Informe de Retiro de Mercaderia";
			this.mnuDTIRM.Id = 686;
			this.mnuDTIRM.Name = "mnuDTIRM";
			this.mnuDTIRM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuDTIRM_ItemClick);
			// 
			// mnuDTIRMP
			// 
			this.mnuDTIRMP.Caption = "Informe de Retiro de Mercaderia Por Articulos";
			this.mnuDTIRMP.Id = 687;
			this.mnuDTIRMP.Name = "mnuDTIRMP";
			this.mnuDTIRMP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuDTIRMP_ItemClick);
			// 
			// mnuDTProductosDeFeria
			// 
			this.mnuDTProductosDeFeria.Caption = "Productos de Feria";
			this.mnuDTProductosDeFeria.Id = 766;
			this.mnuDTProductosDeFeria.Name = "mnuDTProductosDeFeria";
			this.mnuDTProductosDeFeria.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuDTProductosDeFeria_ItemClick);
			// 
			// mnuDTIRAT
			// 
			this.mnuDTIRAT.Caption = "Asignación de Técnicos";
			this.mnuDTIRAT.Id = 736;
			this.mnuDTIRAT.Name = "mnuDTIRAT";
			this.mnuDTIRAT.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem51_ItemClick);
			// 
			// mnuDTIRSNC
			// 
			this.mnuDTIRSNC.Caption = "Solicitud de Notas de Crédito ";
			this.mnuDTIRSNC.Id = 755;
			this.mnuDTIRSNC.Name = "mnuDTIRSNC";
			this.mnuDTIRSNC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem60_ItemClick);
			// 
			// mnuDTAprobacionRemates
			// 
			this.mnuDTAprobacionRemates.Caption = "Envio a Remates";
			this.mnuDTAprobacionRemates.Id = 805;
			this.mnuDTAprobacionRemates.Name = "mnuDTAprobacionRemates";
			this.mnuDTAprobacionRemates.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuDTAprobacionRemates_ItemClick);
			// 
			// mnuDTIRRPT
			// 
			this.mnuDTIRRPT.Caption = "Informes de Reparación";
			this.mnuDTIRRPT.Id = 738;
			this.mnuDTIRRPT.Name = "mnuDTIRRPT";
			this.mnuDTIRRPT.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem53_ItemClick);
			// 
			// mnuDTIRRMA
			// 
			this.mnuDTIRRMA.Caption = "RMA Garantia de Proveedor";
			this.mnuDTIRRMA.Id = 758;
			this.mnuDTIRRMA.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRRPG),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRRSR),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuDTSTA, true)});
			this.mnuDTIRRMA.Name = "mnuDTIRRMA";
			// 
			// mnuDTIRRPG
			// 
			this.mnuDTIRRPG.Caption = "Genera Salida de Articulos";
			this.mnuDTIRRPG.Id = 761;
			this.mnuDTIRRPG.Name = "mnuDTIRRPG";
			this.mnuDTIRRPG.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem64_ItemClick);
			// 
			// mnuDTIRRSR
			// 
			this.mnuDTIRRSR.Caption = "Seguimiento";
			this.mnuDTIRRSR.Id = 762;
			this.mnuDTIRRSR.Name = "mnuDTIRRSR";
			this.mnuDTIRRSR.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem65_ItemClick);
			// 
			// MnuDTSTA
			// 
			this.MnuDTSTA.Caption = "Servicio Técnico Autorizado";
			this.MnuDTSTA.Id = 521;
			this.MnuDTSTA.Name = "MnuDTSTA";
			this.MnuDTSTA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuDTSTA_ItemClick);
			// 
			// frmDTTransfLocales
			// 
			this.frmDTTransfLocales.Caption = "Transferencia a Locales";
			this.frmDTTransfLocales.Id = 800;
			this.frmDTTransfLocales.Name = "frmDTTransfLocales";
			this.frmDTTransfLocales.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.frmDTTransfLocales_ItemClick);
			// 
			// mnuDTIRAPD
			// 
			this.mnuDTIRAPD.Caption = "Autorizacion de Repuestos y Notas de Credito";
			this.mnuDTIRAPD.Id = 754;
			this.mnuDTIRAPD.Name = "mnuDTIRAPD";
			this.mnuDTIRAPD.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem59_ItemClick);
			// 
			// mnuDTASSF
			// 
			this.mnuDTASSF.Caption = "Aprobación Salida Sin Facturar";
			this.mnuDTASSF.Id = 806;
			this.mnuDTASSF.Name = "mnuDTASSF";
			this.mnuDTASSF.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuDTASSF_ItemClick);
			// 
			// mnuDTIRREP
			// 
			this.mnuDTIRREP.Caption = "Reportes";
			this.mnuDTIRREP.Id = 757;
			this.mnuDTIRREP.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTReporteEstadoIngresos),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRRRP),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRRTE),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRFDT),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTIRSIN),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuDTSubGrupo)});
			this.mnuDTIRREP.Name = "mnuDTIRREP";
			// 
			// mnuDTReporteEstadoIngresos
			// 
			this.mnuDTReporteEstadoIngresos.Caption = "Estado Ingresos";
			this.mnuDTReporteEstadoIngresos.Id = 802;
			this.mnuDTReporteEstadoIngresos.Name = "mnuDTReporteEstadoIngresos";
			this.mnuDTReporteEstadoIngresos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuDTReporteEstadoIngresos_ItemClick);
			// 
			// mnuDTIRRRP
			// 
			this.mnuDTIRRRP.Caption = "Reportes de Productos";
			this.mnuDTIRRRP.Id = 759;
			this.mnuDTIRRRP.Name = "mnuDTIRRRP";
			this.mnuDTIRRRP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem62_ItemClick);
			// 
			// mnuDTIRRTE
			// 
			this.mnuDTIRRTE.Caption = "Reportes Técnicos";
			this.mnuDTIRRTE.Id = 760;
			this.mnuDTIRRTE.Name = "mnuDTIRRTE";
			this.mnuDTIRRTE.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem63_ItemClick);
			// 
			// mnuDTIRFDT
			// 
			this.mnuDTIRFDT.Caption = "Facturación";
			this.mnuDTIRFDT.Id = 792;
			this.mnuDTIRFDT.Name = "mnuDTIRFDT";
			this.mnuDTIRFDT.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem73_ItemClick);
			// 
			// mnuDTIRSIN
			// 
			this.mnuDTIRSIN.Caption = "Seguimiento INGRESOS";
			this.mnuDTIRSIN.Id = 793;
			this.mnuDTIRSIN.Name = "mnuDTIRSIN";
			this.mnuDTIRSIN.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem77_ItemClick);
			// 
			// mnuDTSubGrupo
			// 
			this.mnuDTSubGrupo.Caption = "DT SubGrupo";
			this.mnuDTSubGrupo.Id = 849;
			this.mnuDTSubGrupo.Name = "mnuDTSubGrupo";
			this.mnuDTSubGrupo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuDTSubGrupo_ItemClick);
			// 
			// mnuDTIRCDE
			// 
			this.mnuDTIRCDE.Caption = "Cambios de Estado Productos";
			this.mnuDTIRCDE.Id = 794;
			this.mnuDTIRCDE.Name = "mnuDTIRCDE";
			this.mnuDTIRCDE.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem78_ItemClick);
			// 
			// mnuLotte
			// 
			this.mnuLotte.Caption = "Peony &Servicios";
			this.mnuLotte.Id = 637;
			this.mnuLotte.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuCuadroHabitaciones),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem26),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuPSRDC),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuCobrPrenda),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuReservas, true),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuPromociones, true),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuKardexLot, true),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem27),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem28),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem29),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuLoteNotasDeEntrega),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuHabitaciones, true),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuTmaHabitacion),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuTmpEstadia),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuTmpHabitaciones),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem36, true),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.barSubItem9, true),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem30, true),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuLottePlanificacion),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem73),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuLotteCompetencia),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuTarjetasDeIngreso)});
			this.mnuLotte.Name = "mnuLotte";
			// 
			// mnuCuadroHabitaciones
			// 
			this.mnuCuadroHabitaciones.Caption = "Cuadricula Habitaciones";
			this.mnuCuadroHabitaciones.Id = 642;
			this.mnuCuadroHabitaciones.Name = "mnuCuadroHabitaciones";
			this.mnuCuadroHabitaciones.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem33_ItemClick);
			// 
			// barButtonItem26
			// 
			this.barButtonItem26.Caption = "Facturación";
			this.barButtonItem26.Id = 656;
			this.barButtonItem26.Name = "barButtonItem26";
			this.barButtonItem26.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem26_ItemClick);
			// 
			// mnuPSRDC
			// 
			this.mnuPSRDC.Caption = "Resumen Diario de Caja";
			this.mnuPSRDC.Id = 836;
			this.mnuPSRDC.Name = "mnuPSRDC";
			this.mnuPSRDC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuPSRDC_ItemClick);
			// 
			// mnuCobrPrenda
			// 
			this.mnuCobrPrenda.Caption = "Cobro de Prendas";
			this.mnuCobrPrenda.Id = 643;
			this.mnuCobrPrenda.Name = "mnuCobrPrenda";
			this.mnuCobrPrenda.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem34_ItemClick);
			// 
			// mnuReservas
			// 
			this.mnuReservas.Caption = "Reservas";
			this.mnuReservas.Id = 646;
			this.mnuReservas.Name = "mnuReservas";
			this.mnuReservas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuReservas_ItemClick);
			// 
			// mnuPromociones
			// 
			this.mnuPromociones.Caption = "Promociones";
			this.mnuPromociones.Id = 645;
			this.mnuPromociones.Name = "mnuPromociones";
			this.mnuPromociones.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem36_ItemClick);
			// 
			// mnuKardexLot
			// 
			this.mnuKardexLot.Caption = "Kardex";
			this.mnuKardexLot.Id = 644;
			this.mnuKardexLot.Name = "mnuKardexLot";
			this.mnuKardexLot.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem35_ItemClick);
			// 
			// barButtonItem27
			// 
			this.barButtonItem27.Caption = "Minimos de Inventario";
			this.barButtonItem27.Id = 658;
			this.barButtonItem27.Name = "barButtonItem27";
			this.barButtonItem27.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem27_ItemClick);
			// 
			// barButtonItem28
			// 
			this.barButtonItem28.Caption = "Ajustes INVENTARIOS";
			this.barButtonItem28.Id = 661;
			this.barButtonItem28.Name = "barButtonItem28";
			this.barButtonItem28.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem28_ItemClick);
			// 
			// barButtonItem29
			// 
			this.barButtonItem29.Caption = "Transferencia de Bodegas";
			this.barButtonItem29.Id = 662;
			this.barButtonItem29.Name = "barButtonItem29";
			this.barButtonItem29.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem29_ItemClick);
			// 
			// mnuLoteNotasDeEntrega
			// 
			this.mnuLoteNotasDeEntrega.Caption = "Notas De Entrega";
			this.mnuLoteNotasDeEntrega.Id = 669;
			this.mnuLoteNotasDeEntrega.Name = "mnuLoteNotasDeEntrega";
			this.mnuLoteNotasDeEntrega.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuLoteNotasDeEntrega_ItemClick);
			// 
			// mnuHabitaciones
			// 
			this.mnuHabitaciones.Caption = "Habitaciones";
			this.mnuHabitaciones.Id = 638;
			this.mnuHabitaciones.Name = "mnuHabitaciones";
			this.mnuHabitaciones.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem29_ItemClick_1);
			// 
			// mnuTmaHabitacion
			// 
			this.mnuTmaHabitacion.Caption = "Temáticas Habitaciones";
			this.mnuTmaHabitacion.Id = 639;
			this.mnuTmaHabitacion.Name = "mnuTmaHabitacion";
			this.mnuTmaHabitacion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem30_ItemClick);
			// 
			// mnuTmpEstadia
			// 
			this.mnuTmpEstadia.Caption = "Tiempos de Estadía";
			this.mnuTmpEstadia.Id = 640;
			this.mnuTmpEstadia.Name = "mnuTmpEstadia";
			this.mnuTmpEstadia.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem31_ItemClick);
			// 
			// mnuTmpHabitaciones
			// 
			this.mnuTmpHabitaciones.Caption = "Tiempos Habitaciones";
			this.mnuTmpHabitaciones.Id = 641;
			this.mnuTmpHabitaciones.Name = "mnuTmpHabitaciones";
			this.mnuTmpHabitaciones.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem32_ItemClick);
			// 
			// barButtonItem36
			// 
			this.barButtonItem36.Caption = "Porcentaje DESC";
			this.barButtonItem36.Id = 680;
			this.barButtonItem36.Name = "barButtonItem36";
			this.barButtonItem36.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem36_ItemClick_1);
			// 
			// barSubItem9
			// 
			this.barSubItem9.Caption = "Reportes";
			this.barSubItem9.Id = 653;
			this.barSubItem9.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem24),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem25),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem31),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem32),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem22),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem35),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem40),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem42),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem67),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.barSubItem14),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem75),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuLotteReportesIT),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuLotteReporteIPH)});
			this.barSubItem9.Name = "barSubItem9";
			// 
			// barButtonItem24
			// 
			this.barButtonItem24.Caption = "Movimiento Inventario";
			this.barButtonItem24.Id = 654;
			this.barButtonItem24.Name = "barButtonItem24";
			this.barButtonItem24.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem24_ItemClick);
			// 
			// barButtonItem25
			// 
			this.barButtonItem25.Caption = "Caja";
			this.barButtonItem25.Id = 655;
			this.barButtonItem25.Name = "barButtonItem25";
			this.barButtonItem25.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem25_ItemClick);
			// 
			// barButtonItem31
			// 
			this.barButtonItem31.Caption = "Ventas";
			this.barButtonItem31.Id = 665;
			this.barButtonItem31.Name = "barButtonItem31";
			this.barButtonItem31.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem31_ItemClick_1);
			// 
			// barButtonItem32
			// 
			this.barButtonItem32.Caption = "Resumen de Ventas";
			this.barButtonItem32.Id = 666;
			this.barButtonItem32.Name = "barButtonItem32";
			this.barButtonItem32.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem32_ItemClick_1);
			// 
			// barButtonItem22
			// 
			this.barButtonItem22.Caption = "Estado De Reservas";
			this.barButtonItem22.Id = 651;
			this.barButtonItem22.Name = "barButtonItem22";
			this.barButtonItem22.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem22_ItemClick);
			// 
			// barButtonItem35
			// 
			this.barButtonItem35.Caption = "caja gastos";
			this.barButtonItem35.Id = 693;
			this.barButtonItem35.Name = "barButtonItem35";
			this.barButtonItem35.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem35_ItemClick_2);
			// 
			// barButtonItem40
			// 
			this.barButtonItem40.Caption = "venta de articulos";
			this.barButtonItem40.Id = 700;
			this.barButtonItem40.Name = "barButtonItem40";
			this.barButtonItem40.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem40_ItemClick);
			// 
			// barButtonItem42
			// 
			this.barButtonItem42.Caption = "Resumen Diario de Ventas";
			this.barButtonItem42.Id = 710;
			this.barButtonItem42.Name = "barButtonItem42";
			this.barButtonItem42.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem42_ItemClick);
			// 
			// barButtonItem67
			// 
			this.barButtonItem67.Caption = "Venta de Platos por Mes";
			this.barButtonItem67.Id = 765;
			this.barButtonItem67.Name = "barButtonItem67";
			this.barButtonItem67.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem67_ItemClick);
			// 
			// barSubItem14
			// 
			this.barSubItem14.Caption = "Procesos";
			this.barSubItem14.Id = 782;
			this.barSubItem14.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuLoteRepProcSIM),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem69)});
			this.barSubItem14.Name = "barSubItem14";
			// 
			// mnuLoteRepProcSIM
			// 
			this.mnuLoteRepProcSIM.Caption = "Sube Informacion Mensual";
			this.mnuLoteRepProcSIM.Id = 783;
			this.mnuLoteRepProcSIM.Name = "mnuLoteRepProcSIM";
			this.mnuLoteRepProcSIM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuLoteRepProcSIM_ItemClick);
			// 
			// barButtonItem69
			// 
			this.barButtonItem69.Caption = "Calculo de metas mes y año";
			this.barButtonItem69.Id = 768;
			this.barButtonItem69.Name = "barButtonItem69";
			this.barButtonItem69.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem69_ItemClick);
			// 
			// barButtonItem75
			// 
			this.barButtonItem75.Caption = "Compara Datos";
			this.barButtonItem75.Id = 785;
			this.barButtonItem75.Name = "barButtonItem75";
			this.barButtonItem75.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem75_ItemClick);
			// 
			// mnuLotteReportesIT
			// 
			this.mnuLotteReportesIT.Caption = "Ingresos Por Turnos";
			this.mnuLotteReportesIT.Id = 827;
			this.mnuLotteReportesIT.Name = "mnuLotteReportesIT";
			this.mnuLotteReportesIT.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuLotteReportesIT_ItemClick);
			// 
			// mnuLotteReporteIPH
			// 
			this.mnuLotteReporteIPH.Caption = "Ingresos Por Horas";
			this.mnuLotteReporteIPH.Id = 830;
			this.mnuLotteReporteIPH.Name = "mnuLotteReporteIPH";
			this.mnuLotteReporteIPH.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuLotteReporteIPH_ItemClick);
			// 
			// barButtonItem30
			// 
			this.barButtonItem30.Caption = "Turnos Personal";
			this.barButtonItem30.Id = 663;
			this.barButtonItem30.Name = "barButtonItem30";
			this.barButtonItem30.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem30_ItemClick_1);
			// 
			// mnuLottePlanificacion
			// 
			this.mnuLottePlanificacion.Caption = "Planificacion Personal";
			this.mnuLottePlanificacion.Id = 824;
			this.mnuLottePlanificacion.Name = "mnuLottePlanificacion";
			this.mnuLottePlanificacion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuLottePlanificacion_ItemClick);
			// 
			// barButtonItem73
			// 
			this.barButtonItem73.Caption = "Registro de Metas";
			this.barButtonItem73.Id = 828;
			this.barButtonItem73.Name = "barButtonItem73";
			this.barButtonItem73.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem73_ItemClick_1);
			// 
			// mnuLotteCompetencia
			// 
			this.mnuLotteCompetencia.Caption = "Competencia Diaria";
			this.mnuLotteCompetencia.Id = 829;
			this.mnuLotteCompetencia.Name = "mnuLotteCompetencia";
			this.mnuLotteCompetencia.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuLotteCompetencia_ItemClick);
			// 
			// mnuTarjetasDeIngreso
			// 
			this.mnuTarjetasDeIngreso.Caption = "Tarjetas De Ingreso";
			this.mnuTarjetasDeIngreso.Id = 861;
			this.mnuTarjetasDeIngreso.Name = "mnuTarjetasDeIngreso";
			this.mnuTarjetasDeIngreso.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuTarjetasDeIngreso_ItemClick);
			// 
			// MnuVentana
			// 
			this.MnuVentana.Caption = "&Ventana";
			this.MnuVentana.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.MnuVentana.Id = 177;
			this.MnuVentana.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuVentana_Casc),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuVentana_Org),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.MnuVentana_Cicl),
																																														new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.MnuVentTrans, false),
																																														new DevExpress.XtraBars.LinkPersistInfo(this.mnuVentOpac, true)});
			this.MnuVentana.Name = "MnuVentana";
			// 
			// MnuVentana_Casc
			// 
			this.MnuVentana_Casc.Caption = "Cascada";
			this.MnuVentana_Casc.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.MnuVentana_Casc.Id = 191;
			this.MnuVentana_Casc.Name = "MnuVentana_Casc";
			this.MnuVentana_Casc.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuVentana_Casc_ItemClick);
			// 
			// MnuVentana_Org
			// 
			this.MnuVentana_Org.Caption = "Horizontal";
			this.MnuVentana_Org.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.MnuVentana_Org.Id = 192;
			this.MnuVentana_Org.Name = "MnuVentana_Org";
			this.MnuVentana_Org.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuVentana_Org_ItemClick);
			// 
			// MnuVentana_Cicl
			// 
			this.MnuVentana_Cicl.Caption = "Vertical";
			this.MnuVentana_Cicl.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.MnuVentana_Cicl.Id = 193;
			this.MnuVentana_Cicl.Name = "MnuVentana_Cicl";
			this.MnuVentana_Cicl.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuVentana_Cicl_ItemClick);
			// 
			// MnuVentTrans
			// 
			this.MnuVentTrans.Caption = "Transparente";
			this.MnuVentTrans.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.MnuVentTrans.Id = 262;
			this.MnuVentTrans.Name = "MnuVentTrans";
			this.MnuVentTrans.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuVentTrans_ItemClick);
			// 
			// mnuVentOpac
			// 
			this.mnuVentOpac.Caption = "Opacidad";
			this.mnuVentOpac.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.mnuVentOpac.Id = 263;
			this.mnuVentOpac.Name = "mnuVentOpac";
			this.mnuVentOpac.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentOpac_ItemClick);
			// 
			// MnuFact
			// 
			this.MnuFact.Caption = "&Facturación";
			this.MnuFact.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFact.Id = 104;
			this.MnuFact.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactCotizacion),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactResPed, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactLotes, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactGrupArticulos, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFacSerie, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactCtasCobrar, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactCtasPagar),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactRepResumenes, true),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactRepInventario),
																																												 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactRepArticulo)});
			this.MnuFact.Name = "MnuFact";
			// 
			// MnuFactCotizacion
			// 
			this.MnuFactCotizacion.Caption = "Cotización";
			this.MnuFactCotizacion.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactCotizacion.Id = 108;
			this.MnuFactCotizacion.Name = "MnuFactCotizacion";
			this.MnuFactCotizacion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactCotizacion_ItemClick);
			// 
			// MnuFactResPed
			// 
			this.MnuFactResPed.Caption = "Reservas y Pedidos";
			this.MnuFactResPed.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactResPed.Id = 204;
			this.MnuFactResPed.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactReservProveedores),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactResPedProv),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuResPedPend),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactPreciosProveedores),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactResPedRequisicion),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactResPedComision),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactResCruzarFactReserva),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactResOTrabajo),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactResConclusion),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactResOTrabajoServicio),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactResJarComision)});
			this.MnuFactResPed.Name = "MnuFactResPed";
			// 
			// MnuFactReservProveedores
			// 
			this.MnuFactReservProveedores.Caption = "Reservación Proveedores";
			this.MnuFactReservProveedores.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactReservProveedores.Id = 114;
			this.MnuFactReservProveedores.Name = "MnuFactReservProveedores";
			this.MnuFactReservProveedores.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactReservProveedores_ItemClick);
			// 
			// MnuFactResPedProv
			// 
			this.MnuFactResPedProv.Caption = "Pedido Proveedor";
			this.MnuFactResPedProv.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactResPedProv.Id = 270;
			this.MnuFactResPedProv.Name = "MnuFactResPedProv";
			this.MnuFactResPedProv.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactResPedProv_ItemClick);
			// 
			// mnuResPedPend
			// 
			this.mnuResPedPend.Caption = "Pedidos Pendientes";
			this.mnuResPedPend.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.mnuResPedPend.Id = 264;
			this.mnuResPedPend.Name = "mnuResPedPend";
			this.mnuResPedPend.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactPedPend);
			// 
			// MnuFactPreciosProveedores
			// 
			this.MnuFactPreciosProveedores.Caption = "Precios de Proveedores";
			this.MnuFactPreciosProveedores.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactPreciosProveedores.Id = 117;
			this.MnuFactPreciosProveedores.Name = "MnuFactPreciosProveedores";
			// 
			// MnuFactResPedRequisicion
			// 
			this.MnuFactResPedRequisicion.Caption = "Requisición";
			this.MnuFactResPedRequisicion.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactResPedRequisicion.Id = 383;
			this.MnuFactResPedRequisicion.Name = "MnuFactResPedRequisicion";
			this.MnuFactResPedRequisicion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btRequisicion_ItemClick);
			// 
			// MnuFactResPedComision
			// 
			this.MnuFactResPedComision.Caption = "Comisión";
			this.MnuFactResPedComision.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactResPedComision.Id = 385;
			this.MnuFactResPedComision.Name = "MnuFactResPedComision";
			this.MnuFactResPedComision.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactResPedComision_ItemClick);
			// 
			// MnuFactResCruzarFactReserva
			// 
			this.MnuFactResCruzarFactReserva.Caption = "Cruzar Reserva/Factura";
			this.MnuFactResCruzarFactReserva.Id = 423;
			this.MnuFactResCruzarFactReserva.Name = "MnuFactResCruzarFactReserva";
			this.MnuFactResCruzarFactReserva.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactResCruzarFactReserva_ItemClick);
			// 
			// MnuFactResOTrabajo
			// 
			this.MnuFactResOTrabajo.Caption = "Orden Trabajo";
			this.MnuFactResOTrabajo.Id = 426;
			this.MnuFactResOTrabajo.Name = "MnuFactResOTrabajo";
			this.MnuFactResOTrabajo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactResOTrabajo_ItemClick);
			// 
			// MnuFactResConclusion
			// 
			this.MnuFactResConclusion.Caption = "Conclusión";
			this.MnuFactResConclusion.Id = 427;
			this.MnuFactResConclusion.Name = "MnuFactResConclusion";
			this.MnuFactResConclusion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactResConclusion_ItemClick);
			// 
			// MnuFactResOTrabajoServicio
			// 
			this.MnuFactResOTrabajoServicio.Caption = "Orden Trabajo Servicio";
			this.MnuFactResOTrabajoServicio.Id = 428;
			this.MnuFactResOTrabajoServicio.Name = "MnuFactResOTrabajoServicio";
			this.MnuFactResOTrabajoServicio.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactResOTrabajoServicio_ItemClick);
			// 
			// MnuFactResJarComision
			// 
			this.MnuFactResJarComision.Caption = "Comisiones Jardines";
			this.MnuFactResJarComision.Id = 436;
			this.MnuFactResJarComision.Name = "MnuFactResJarComision";
			this.MnuFactResJarComision.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactResJarComision_ItemClick);
			// 
			// MnuFactLotes
			// 
			this.MnuFactLotes.Caption = "Lotes";
			this.MnuFactLotes.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactLotes.Id = 118;
			this.MnuFactLotes.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																															new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactLot_Ctas_Cob),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactLot_Ctas_Pag),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactLot_Autoriza),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactLot_Facturacion, true),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuFactLotesPago),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactLot_Transf_Art),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuFacLotTransfTrans),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.MnuFactLoteAnalisis),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuFactLotAgencia),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuFactLotTomaF),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuFacLotCxC30),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuFactLoteTarjeta),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuFacLotSerie),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuFactRepSri),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuFacLotCierre),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuFacLotCopiaArt),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuFacLotFactReembol)});
			this.MnuFactLotes.Name = "MnuFactLotes";
			this.MnuFactLotes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// MnuFactLot_Ctas_Cob
			// 
			this.MnuFactLot_Ctas_Cob.Caption = "Cuentas por Cobrar";
			this.MnuFactLot_Ctas_Cob.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactLot_Ctas_Cob.Id = 119;
			this.MnuFactLot_Ctas_Cob.Name = "MnuFactLot_Ctas_Cob";
			this.MnuFactLot_Ctas_Cob.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactLot_Ctas_Cob_ItemClick);
			// 
			// MnuFactLot_Ctas_Pag
			// 
			this.MnuFactLot_Ctas_Pag.Caption = "Cuentas por Pagar";
			this.MnuFactLot_Ctas_Pag.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactLot_Ctas_Pag.Id = 120;
			this.MnuFactLot_Ctas_Pag.Name = "MnuFactLot_Ctas_Pag";
			this.MnuFactLot_Ctas_Pag.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactLot_Ctas_Pag_ItemClick);
			// 
			// MnuFactLot_Autoriza
			// 
			this.MnuFactLot_Autoriza.Caption = "Autorización de Pagos";
			this.MnuFactLot_Autoriza.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactLot_Autoriza.Id = 289;
			this.MnuFactLot_Autoriza.Name = "MnuFactLot_Autoriza";
			this.MnuFactLot_Autoriza.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactLot_Autoriza_ItemClick);
			// 
			// MnuFactLot_Facturacion
			// 
			this.MnuFactLot_Facturacion.Caption = "Facturación";
			this.MnuFactLot_Facturacion.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactLot_Facturacion.Id = 121;
			this.MnuFactLot_Facturacion.Name = "MnuFactLot_Facturacion";
			this.MnuFactLot_Facturacion.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactLot_Facturacion_ItemClick);
			// 
			// mnuFactLotesPago
			// 
			this.mnuFactLotesPago.Caption = "Depósitos";
			this.mnuFactLotesPago.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.mnuFactLotesPago.Id = 267;
			this.mnuFactLotesPago.Name = "mnuFactLotesPago";
			this.mnuFactLotesPago.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFactLotesPago_ItemClick);
			// 
			// MnuFactLot_Transf_Art
			// 
			this.MnuFactLot_Transf_Art.Caption = "Transferencia de Artículos";
			this.MnuFactLot_Transf_Art.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactLot_Transf_Art.Id = 124;
			this.MnuFactLot_Transf_Art.Name = "MnuFactLot_Transf_Art";
			this.MnuFactLot_Transf_Art.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactLot_Transf_Art_ItemClick);
			// 
			// mnuFacLotTransfTrans
			// 
			this.mnuFacLotTransfTrans.Caption = "Transferencia Transacción";
			this.mnuFacLotTransfTrans.Id = 410;
			this.mnuFacLotTransfTrans.Name = "mnuFacLotTransfTrans";
			this.mnuFacLotTransfTrans.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFacLotTransfTrans_ItemClick);
			// 
			// MnuFactLoteAnalisis
			// 
			this.MnuFactLoteAnalisis.Caption = "Análisis de Artículos";
			this.MnuFactLoteAnalisis.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactLoteAnalisis.Id = 249;
			this.MnuFactLoteAnalisis.Name = "MnuFactLoteAnalisis";
			this.MnuFactLoteAnalisis.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactLoteAnalisis_ItemClick);
			// 
			// mnuFactLotAgencia
			// 
			this.mnuFactLotAgencia.Caption = "Agencia de Viaje";
			this.mnuFactLotAgencia.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.mnuFactLotAgencia.Id = 254;
			this.mnuFactLotAgencia.Name = "mnuFactLotAgencia";
			this.mnuFactLotAgencia.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFactLotAgencia_ItemClick);
			// 
			// mnuFactLotTomaF
			// 
			this.mnuFactLotTomaF.Caption = "Toma Física";
			this.mnuFactLotTomaF.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.mnuFactLotTomaF.Id = 274;
			this.mnuFactLotTomaF.Name = "mnuFactLotTomaF";
			this.mnuFactLotTomaF.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFactLotTomaF_ItemClick);
			// 
			// mnuFacLotCxC30
			// 
			this.mnuFacLotCxC30.Caption = "CxC y P 30 60 90";
			this.mnuFacLotCxC30.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.mnuFacLotCxC30.Id = 276;
			this.mnuFacLotCxC30.Name = "mnuFacLotCxC30";
			this.mnuFacLotCxC30.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFacLotCxC30_ItemClick);
			// 
			// mnuFactLoteTarjeta
			// 
			this.mnuFactLoteTarjeta.Caption = "Tarjeta de Crédito";
			this.mnuFactLoteTarjeta.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.mnuFactLoteTarjeta.Id = 283;
			this.mnuFactLoteTarjeta.Name = "mnuFactLoteTarjeta";
			this.mnuFactLoteTarjeta.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFactLoteTarjeta_ItemClick);
			// 
			// mnuFacLotSerie
			// 
			this.mnuFacLotSerie.Caption = "Serie";
			this.mnuFacLotSerie.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.mnuFacLotSerie.Id = 284;
			this.mnuFacLotSerie.Name = "mnuFacLotSerie";
			this.mnuFacLotSerie.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFacLotSerie_ItemClick);
			// 
			// mnuFactRepSri
			// 
			this.mnuFactRepSri.Caption = "Impresión SRI";
			this.mnuFactRepSri.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.mnuFactRepSri.Id = 298;
			this.mnuFactRepSri.Name = "mnuFactRepSri";
			this.mnuFactRepSri.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFacRepSri);
			// 
			// mnuFacLotCierre
			// 
			this.mnuFacLotCierre.Caption = "Cierre de Caja";
			this.mnuFacLotCierre.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.mnuFacLotCierre.Id = 396;
			this.mnuFacLotCierre.Name = "mnuFacLotCierre";
			this.mnuFacLotCierre.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFacLotCierre_ItemClick);
			// 
			// mnuFacLotCopiaArt
			// 
			this.mnuFacLotCopiaArt.Caption = "Crea Artículo Compra";
			this.mnuFacLotCopiaArt.Id = 415;
			this.mnuFacLotCopiaArt.Name = "mnuFacLotCopiaArt";
			this.mnuFacLotCopiaArt.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFacLotCopiaArt_ItemClick);
			// 
			// mnuFacLotFactReembol
			// 
			this.mnuFacLotFactReembol.Caption = "Genera Factura/Reembolso";
			this.mnuFacLotFactReembol.Id = 418;
			this.mnuFacLotFactReembol.Name = "mnuFacLotFactReembol";
			this.mnuFacLotFactReembol.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuFacLotFactReembol_ItemClick);
			// 
			// MnuFactGrupArticulos
			// 
			this.MnuFactGrupArticulos.Caption = "Grupo de Artículos";
			this.MnuFactGrupArticulos.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactGrupArticulos.Id = 128;
			this.MnuFactGrupArticulos.Name = "MnuFactGrupArticulos";
			this.MnuFactGrupArticulos.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactGrupArticulos_ItemClick);
			// 
			// MnuFacSerie
			// 
			this.MnuFacSerie.Caption = "&Dep.Tecnico";
			this.MnuFacSerie.Id = 461;
			this.MnuFacSerie.Name = "MnuFacSerie";
			this.MnuFacSerie.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.MnuFacSerie.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFacSerie_ItemClick);
			// 
			// MnuFactCtasCobrar
			// 
			this.MnuFactCtasCobrar.Caption = "Cuentas por Cobrar";
			this.MnuFactCtasCobrar.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactCtasCobrar.Id = 131;
			this.MnuFactCtasCobrar.ImageIndex = 47;
			this.MnuFactCtasCobrar.ItemShortcut = new DevExpress.XtraBars.BarShortcut(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
				| System.Windows.Forms.Keys.C));
			this.MnuFactCtasCobrar.Name = "MnuFactCtasCobrar";
			this.MnuFactCtasCobrar.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactCtasCobrar_ItemClick);
			// 
			// MnuFactCtasPagar
			// 
			this.MnuFactCtasPagar.Caption = "Cuentas por Pagar  ";
			this.MnuFactCtasPagar.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactCtasPagar.Id = 132;
			this.MnuFactCtasPagar.ImageIndex = 45;
			this.MnuFactCtasPagar.ItemShortcut = new DevExpress.XtraBars.BarShortcut(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
				| System.Windows.Forms.Keys.P));
			this.MnuFactCtasPagar.Name = "MnuFactCtasPagar";
			this.MnuFactCtasPagar.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactCtasPagar_ItemClick);
			// 
			// MnuFactRepResumenes
			// 
			this.MnuFactRepResumenes.Caption = "Reporte Resúmenes         ";
			this.MnuFactRepResumenes.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactRepResumenes.Id = 133;
			this.MnuFactRepResumenes.ImageIndex = 24;
			this.MnuFactRepResumenes.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.D));
			this.MnuFactRepResumenes.Name = "MnuFactRepResumenes";
			this.MnuFactRepResumenes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactRepResumenes_ItemClick);
			// 
			// MnuFactRepInventario
			// 
			this.MnuFactRepInventario.Caption = "Reporte Inventario";
			this.MnuFactRepInventario.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactRepInventario.Id = 134;
			this.MnuFactRepInventario.ImageIndex = 23;
			this.MnuFactRepInventario.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.I));
			this.MnuFactRepInventario.Name = "MnuFactRepInventario";
			this.MnuFactRepInventario.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactRepInventario_ItemClick);
			// 
			// MnuFactRepArticulo
			// 
			this.MnuFactRepArticulo.Caption = "Reporte de Artículos";
			this.MnuFactRepArticulo.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactRepArticulo.Id = 135;
			this.MnuFactRepArticulo.ImageIndex = 25;
			this.MnuFactRepArticulo.Name = "MnuFactRepArticulo";
			this.MnuFactRepArticulo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactGraficos_ItemClick);
			// 
			// mnuMetricas
			// 
			this.mnuMetricas.Caption = "Metricas";
			this.mnuMetricas.Id = 812;
			this.mnuMetricas.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuMetricasRRHH)});
			this.mnuMetricas.Name = "mnuMetricas";
			// 
			// mnuMetricasRRHH
			// 
			this.mnuMetricasRRHH.Caption = "RRHH";
			this.mnuMetricasRRHH.Id = 814;
			this.mnuMetricasRRHH.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem60),
																																																 new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem62)});
			this.mnuMetricasRRHH.Name = "mnuMetricasRRHH";
			// 
			// barButtonItem60
			// 
			this.barButtonItem60.Caption = "I.R.P.";
			this.barButtonItem60.Id = 815;
			this.barButtonItem60.Name = "barButtonItem60";
			this.barButtonItem60.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem60_ItemClick_1);
			// 
			// barButtonItem62
			// 
			this.barButtonItem62.Caption = "Contratación";
			this.barButtonItem62.Id = 816;
			this.barButtonItem62.Name = "barButtonItem62";
			this.barButtonItem62.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem62_ItemClick_1);
			// 
			// mnuPCGEx
			// 
			this.mnuPCGEx.Caption = "Politicas de Credito y Garantias";
			this.mnuPCGEx.Id = 819;
			this.mnuPCGEx.Name = "mnuPCGEx";
			this.mnuPCGEx.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuPCGEx_ItemClick);
			// 
			// mnuSTA
			// 
			this.mnuSTA.Caption = "Centros de Servicio Tecnico Autorizados";
			this.mnuSTA.Id = 820;
			this.mnuSTA.Name = "mnuSTA";
			this.mnuSTA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuSTA_ItemClick);
			// 
			// mnuDirectorio
			// 
			this.mnuDirectorio.Caption = "Directorio ";
			this.mnuDirectorio.Id = 863;
			this.mnuDirectorio.Name = "mnuDirectorio";
			this.mnuDirectorio.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuDirectorio_ItemClick);
			// 
			// barSubItem5
			// 
			this.barSubItem5.Caption = "Reporteria";
			this.barSubItem5.Id = 898;
			this.barSubItem5.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtBodega),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtContabilidad),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtRoles),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtCompras),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtVentas),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtCREPOINT),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtDT),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtPeonyServicios)});
			this.barSubItem5.Name = "barSubItem5";
			// 
			// mnuRPtBodega
			// 
			this.mnuRPtBodega.Caption = "Bodega";
			this.mnuRPtBodega.Id = 900;
			this.mnuRPtBodega.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtBodegaInvFEcha),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtBodegaRptTrans),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtBodegaInvVal),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtBodegaKXSer),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtBodegaEsConInv),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtBodegaMovDesc),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtBodegaKCC)});
			this.mnuRPtBodega.Name = "mnuRPtBodega";
			// 
			// mnuRPtBodegaInvFEcha
			// 
			this.mnuRPtBodegaInvFEcha.Caption = "Inventario a la Fecha";
			this.mnuRPtBodegaInvFEcha.Id = 912;
			this.mnuRPtBodegaInvFEcha.Name = "mnuRPtBodegaInvFEcha";
			this.mnuRPtBodegaInvFEcha.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtBodegaInvFEcha_ItemClick);
			// 
			// mnuRPtBodegaRptTrans
			// 
			this.mnuRPtBodegaRptTrans.Caption = "Reporte de Transferencia";
			this.mnuRPtBodegaRptTrans.Id = 913;
			this.mnuRPtBodegaRptTrans.Name = "mnuRPtBodegaRptTrans";
			this.mnuRPtBodegaRptTrans.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtBodegaRptTrans_ItemClick);
			// 
			// mnuRPtBodegaInvVal
			// 
			this.mnuRPtBodegaInvVal.Caption = "Inventario Valorado";
			this.mnuRPtBodegaInvVal.Id = 914;
			this.mnuRPtBodegaInvVal.Name = "mnuRPtBodegaInvVal";
			this.mnuRPtBodegaInvVal.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtBodegaInvVal_ItemClick);
			// 
			// mnuRPtBodegaKXSer
			// 
			this.mnuRPtBodegaKXSer.Caption = "Kardex Seriales";
			this.mnuRPtBodegaKXSer.Id = 915;
			this.mnuRPtBodegaKXSer.Name = "mnuRPtBodegaKXSer";
			this.mnuRPtBodegaKXSer.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtBodegaKXSer_ItemClick);
			// 
			// mnuRPtBodegaEsConInv
			// 
			this.mnuRPtBodegaEsConInv.Caption = "Estados de Conteos de Inventarios";
			this.mnuRPtBodegaEsConInv.Id = 916;
			this.mnuRPtBodegaEsConInv.Name = "mnuRPtBodegaEsConInv";
			this.mnuRPtBodegaEsConInv.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtBodegaEsConInv_ItemClick);
			// 
			// mnuRPtBodegaMovDesc
			// 
			this.mnuRPtBodegaMovDesc.Caption = "Movimientos Descuadrados";
			this.mnuRPtBodegaMovDesc.Id = 917;
			this.mnuRPtBodegaMovDesc.Name = "mnuRPtBodegaMovDesc";
			this.mnuRPtBodegaMovDesc.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtBodegaMovDesc_ItemClick);
			// 
			// mnuRPtBodegaKCC
			// 
			this.mnuRPtBodegaKCC.Caption = "Kardex Con Costo";
			this.mnuRPtBodegaKCC.Id = 940;
			this.mnuRPtBodegaKCC.Name = "mnuRPtBodegaKCC";
			this.mnuRPtBodegaKCC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtBodegaKCC_ItemClick);
			// 
			// mnuRPtContabilidad
			// 
			this.mnuRPtContabilidad.Caption = "Contabilidad";
			this.mnuRPtContabilidad.Id = 901;
			this.mnuRPtContabilidad.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtContabilidadCont),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtContabilidadRet),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtContabilidadAsDEsc),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtContabilidadGast),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtContabilidadFactEl),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtContabilidadCaCh),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtContabilidadGaAn),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtContabilidadEstaConB),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtContabilidadPrXAlm),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPPagoAgil),
																																																		new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtContabilidadSCosecha)});
			this.mnuRPtContabilidad.Name = "mnuRPtContabilidad";
			// 
			// mnuRPtContabilidadCont
			// 
			this.mnuRPtContabilidadCont.Caption = "Contabilidad   Ctrl + B";
			this.mnuRPtContabilidadCont.Id = 918;
			this.mnuRPtContabilidadCont.Name = "mnuRPtContabilidadCont";
			this.mnuRPtContabilidadCont.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtContabilidadCont_ItemClick);
			// 
			// mnuRPtContabilidadRet
			// 
			this.mnuRPtContabilidadRet.Caption = "Retenciones";
			this.mnuRPtContabilidadRet.Id = 919;
			this.mnuRPtContabilidadRet.Name = "mnuRPtContabilidadRet";
			this.mnuRPtContabilidadRet.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtContabilidadRet_ItemClick);
			// 
			// mnuRPtContabilidadAsDEsc
			// 
			this.mnuRPtContabilidadAsDEsc.Caption = "Asientos Descuadrados";
			this.mnuRPtContabilidadAsDEsc.Id = 920;
			this.mnuRPtContabilidadAsDEsc.Name = "mnuRPtContabilidadAsDEsc";
			this.mnuRPtContabilidadAsDEsc.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtContabilidadAsDEsc_ItemClick);
			// 
			// mnuRPtContabilidadGast
			// 
			this.mnuRPtContabilidadGast.Caption = "Gastos";
			this.mnuRPtContabilidadGast.Id = 921;
			this.mnuRPtContabilidadGast.Name = "mnuRPtContabilidadGast";
			this.mnuRPtContabilidadGast.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtContabilidadGast_ItemClick);
			// 
			// mnuRPtContabilidadFactEl
			// 
			this.mnuRPtContabilidadFactEl.Caption = "Facturacion Electronica";
			this.mnuRPtContabilidadFactEl.Id = 922;
			this.mnuRPtContabilidadFactEl.Name = "mnuRPtContabilidadFactEl";
			this.mnuRPtContabilidadFactEl.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtContabilidadFactEl_ItemClick);
			// 
			// mnuRPtContabilidadCaCh
			// 
			this.mnuRPtContabilidadCaCh.Caption = "Cajas Chicas";
			this.mnuRPtContabilidadCaCh.Id = 923;
			this.mnuRPtContabilidadCaCh.Name = "mnuRPtContabilidadCaCh";
			this.mnuRPtContabilidadCaCh.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtContabilidadCaCh_ItemClick);
			// 
			// mnuRPtContabilidadGaAn
			// 
			this.mnuRPtContabilidadGaAn.Caption = "Gastos Anuales";
			this.mnuRPtContabilidadGaAn.Id = 924;
			this.mnuRPtContabilidadGaAn.Name = "mnuRPtContabilidadGaAn";
			this.mnuRPtContabilidadGaAn.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtContabilidadGaAn_ItemClick);
			// 
			// mnuRPtContabilidadEstaConB
			// 
			this.mnuRPtContabilidadEstaConB.Caption = "Estado Conciliación Bancaria";
			this.mnuRPtContabilidadEstaConB.Id = 925;
			this.mnuRPtContabilidadEstaConB.Name = "mnuRPtContabilidadEstaConB";
			this.mnuRPtContabilidadEstaConB.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtContabilidadEstaConB_ItemClick);
			// 
			// mnuRPtContabilidadPrXAlm
			// 
			this.mnuRPtContabilidadPrXAlm.Caption = "Presupuesto Por Almacen";
			this.mnuRPtContabilidadPrXAlm.Id = 926;
			this.mnuRPtContabilidadPrXAlm.Name = "mnuRPtContabilidadPrXAlm";
			this.mnuRPtContabilidadPrXAlm.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtContabilidadPrXAlm_ItemClick);
			// 
			// mnuRPPagoAgil
			// 
			this.mnuRPPagoAgil.Caption = "Reportes Pago Agil";
			this.mnuRPPagoAgil.Id = 985;
			this.mnuRPPagoAgil.Name = "mnuRPPagoAgil";
			this.mnuRPPagoAgil.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPPagoAgil_ItemClick);
			// 
			// mnuRPtContabilidadSCosecha
			// 
			this.mnuRPtContabilidadSCosecha.Caption = "Reporte Siembra Cosecha";
			this.mnuRPtContabilidadSCosecha.Id = 1002;
			this.mnuRPtContabilidadSCosecha.Name = "mnuRPtContabilidadSCosecha";
			this.mnuRPtContabilidadSCosecha.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtContabilidadSCosecha_ItemClick);
			// 
			// mnuRPtRoles
			// 
			this.mnuRPtRoles.Caption = "Roles";
			this.mnuRPtRoles.Id = 904;
			this.mnuRPtRoles.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtRolesPFA),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtRolesVeXM),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtRolesUni),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtRolesNG),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtRolesAEM),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtRolesHHT),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtRolesIXD)});
			this.mnuRPtRoles.Name = "mnuRPtRoles";
			// 
			// mnuRPtRolesPFA
			// 
			this.mnuRPtRolesPFA.Caption = "Permisos - Vacaciones -Faltas";
			this.mnuRPtRolesPFA.Id = 927;
			this.mnuRPtRolesPFA.Name = "mnuRPtRolesPFA";
			this.mnuRPtRolesPFA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtRolesPFA_ItemClick);
			// 
			// mnuRPtRolesVeXM
			// 
			this.mnuRPtRolesVeXM.Caption = "Vendedores Por Mes";
			this.mnuRPtRolesVeXM.Id = 928;
			this.mnuRPtRolesVeXM.Name = "mnuRPtRolesVeXM";
			this.mnuRPtRolesVeXM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtRolesVeXM_ItemClick);
			// 
			// mnuRPtRolesUni
			// 
			this.mnuRPtRolesUni.Caption = "Uniformes";
			this.mnuRPtRolesUni.Id = 929;
			this.mnuRPtRolesUni.Name = "mnuRPtRolesUni";
			this.mnuRPtRolesUni.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtRolesUni_ItemClick);
			// 
			// mnuRPtRolesNG
			// 
			this.mnuRPtRolesNG.Caption = "Nomina General";
			this.mnuRPtRolesNG.Id = 930;
			this.mnuRPtRolesNG.Name = "mnuRPtRolesNG";
			this.mnuRPtRolesNG.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtRolesNG_ItemClick);
			// 
			// mnuRPtRolesAEM
			// 
			this.mnuRPtRolesAEM.Caption = "Anticipos Empleados";
			this.mnuRPtRolesAEM.Id = 931;
			this.mnuRPtRolesAEM.Name = "mnuRPtRolesAEM";
			this.mnuRPtRolesAEM.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtRolesAEM_ItemClick);
			// 
			// mnuRPtRolesHHT
			// 
			this.mnuRPtRolesHHT.Caption = "Horas Hombre Trabajadas";
			this.mnuRPtRolesHHT.Id = 932;
			this.mnuRPtRolesHHT.Name = "mnuRPtRolesHHT";
			this.mnuRPtRolesHHT.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtRolesHHT_ItemClick);
			// 
			// mnuRPtRolesIXD
			// 
			this.mnuRPtRolesIXD.Caption = "Ingresos Por Departamento";
			this.mnuRPtRolesIXD.Id = 933;
			this.mnuRPtRolesIXD.Name = "mnuRPtRolesIXD";
			this.mnuRPtRolesIXD.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtRolesIXD_ItemClick);
			// 
			// mnuRPtCompras
			// 
			this.mnuRPtCompras.Caption = "Compras";
			this.mnuRPtCompras.Id = 906;
			this.mnuRPtCompras.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtComprasCPA),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtComprasT),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtComprasRAC),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtComprasCA),
																																															 new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtComprasPuC)});
			this.mnuRPtCompras.Name = "mnuRPtCompras";
			// 
			// mnuRPtComprasCPA
			// 
			this.mnuRPtComprasCPA.Caption = "Compras Por Articulo";
			this.mnuRPtComprasCPA.Id = 934;
			this.mnuRPtComprasCPA.Name = "mnuRPtComprasCPA";
			this.mnuRPtComprasCPA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtComprasCPA_ItemClick);
			// 
			// mnuRPtComprasT
			// 
			this.mnuRPtComprasT.Caption = "Flujo Proveedores";
			this.mnuRPtComprasT.Id = 935;
			this.mnuRPtComprasT.Name = "mnuRPtComprasT";
			this.mnuRPtComprasT.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtComprasT_ItemClick);
			// 
			// mnuRPtComprasRAC
			// 
			this.mnuRPtComprasRAC.Caption = "Articulos a Comprar";
			this.mnuRPtComprasRAC.Id = 936;
			this.mnuRPtComprasRAC.Name = "mnuRPtComprasRAC";
			this.mnuRPtComprasRAC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtComprasRAC_ItemClick);
			// 
			// mnuRPtComprasCA
			// 
			this.mnuRPtComprasCA.Caption = "Codigo Articulo";
			this.mnuRPtComprasCA.Id = 937;
			this.mnuRPtComprasCA.Name = "mnuRPtComprasCA";
			this.mnuRPtComprasCA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtComprasCA_ItemClick);
			// 
			// mnuRPtComprasPuC
			// 
			this.mnuRPtComprasPuC.Caption = "Venta Por Ultimo Costo";
			this.mnuRPtComprasPuC.Id = 987;
			this.mnuRPtComprasPuC.Name = "mnuRPtComprasPuC";
			this.mnuRPtComprasPuC.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtComprasPuC_ItemClick);
			// 
			// mnuRPtVentas
			// 
			this.mnuRPtVentas.Caption = "Ventas";
			this.mnuRPtVentas.Id = 908;
			this.mnuRPtVentas.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasPorcMarca),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasPorcMarcaUtilid, true),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasMargUti),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasReMarUti),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasMaUXArt),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasAnInMEn),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasRevFact, true),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasXDia),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasNFactPForPago),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasGaraExt),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasXFecha),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasComDeVentas),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasANuXLo),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasMenXLo),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasTarCredi),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasDetTc),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasEncuestas, true),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasReCXCo),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasXano),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasXArtic),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasArtXMes),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasPorcMargUtilidad),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.barSubItem12),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuReporVentasPDGA),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuReporVentasDFI)});
			this.mnuRPtVentas.Name = "mnuRPtVentas";
			// 
			// mnuRPTVentasPorcMarca
			// 
			this.mnuRPTVentasPorcMarca.Caption = "Ventas Por Marca";
			this.mnuRPTVentasPorcMarca.Id = 941;
			this.mnuRPTVentasPorcMarca.Name = "mnuRPTVentasPorcMarca";
			this.mnuRPTVentasPorcMarca.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasPorcMarcaUtilid
			// 
			this.mnuRPTVentasPorcMarcaUtilid.Caption = "Ventas Por Marca y Local";
			this.mnuRPTVentasPorcMarcaUtilid.Id = 942;
			this.mnuRPTVentasPorcMarcaUtilid.Name = "mnuRPTVentasPorcMarcaUtilid";
			this.mnuRPTVentasPorcMarcaUtilid.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasMargUti
			// 
			this.mnuRPTVentasMargUti.Caption = "Margen De Utilidad";
			this.mnuRPTVentasMargUti.Id = 943;
			this.mnuRPTVentasMargUti.Name = "mnuRPTVentasMargUti";
			this.mnuRPTVentasMargUti.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasReMarUti
			// 
			this.mnuRPTVentasReMarUti.Caption = "Resumen Margen De Utilidad";
			this.mnuRPTVentasReMarUti.Id = 944;
			this.mnuRPTVentasReMarUti.Name = "mnuRPTVentasReMarUti";
			this.mnuRPTVentasReMarUti.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasMaUXArt
			// 
			this.mnuRPTVentasMaUXArt.Caption = "Margen Utilidad Por Articulo";
			this.mnuRPTVentasMaUXArt.Id = 945;
			this.mnuRPTVentasMaUXArt.Name = "mnuRPTVentasMaUXArt";
			this.mnuRPTVentasMaUXArt.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasAnInMEn
			// 
			this.mnuRPTVentasAnInMEn.Caption = "Analisis Información Mensual";
			this.mnuRPTVentasAnInMEn.Id = 946;
			this.mnuRPTVentasAnInMEn.Name = "mnuRPTVentasAnInMEn";
			this.mnuRPTVentasAnInMEn.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasRevFact
			// 
			this.mnuRPTVentasRevFact.Caption = "Revión de Facturas";
			this.mnuRPTVentasRevFact.Id = 947;
			this.mnuRPTVentasRevFact.Name = "mnuRPTVentasRevFact";
			this.mnuRPTVentasRevFact.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasXDia
			// 
			this.mnuRPTVentasXDia.Caption = "Ventas Por Día";
			this.mnuRPTVentasXDia.Id = 948;
			this.mnuRPTVentasXDia.Name = "mnuRPTVentasXDia";
			this.mnuRPTVentasXDia.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasNFactPForPago
			// 
			this.mnuRPTVentasNFactPForPago.Caption = "Numero De Facturas Por Formas De Pago";
			this.mnuRPTVentasNFactPForPago.Id = 949;
			this.mnuRPTVentasNFactPForPago.Name = "mnuRPTVentasNFactPForPago";
			this.mnuRPTVentasNFactPForPago.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasGaraExt
			// 
			this.mnuRPTVentasGaraExt.Caption = "Garantia Extendida";
			this.mnuRPTVentasGaraExt.Id = 950;
			this.mnuRPTVentasGaraExt.Name = "mnuRPTVentasGaraExt";
			this.mnuRPTVentasGaraExt.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasXFecha
			// 
			this.mnuRPTVentasXFecha.Caption = "Vendedores Por Fecha";
			this.mnuRPTVentasXFecha.Id = 951;
			this.mnuRPTVentasXFecha.Name = "mnuRPTVentasXFecha";
			this.mnuRPTVentasXFecha.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasComDeVentas
			// 
			this.mnuRPTVentasComDeVentas.Caption = "Comparativos de Ventas";
			this.mnuRPTVentasComDeVentas.Id = 952;
			this.mnuRPTVentasComDeVentas.Name = "mnuRPTVentasComDeVentas";
			this.mnuRPTVentasComDeVentas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasANuXLo
			// 
			this.mnuRPTVentasANuXLo.Caption = "Ventas Anuales Por Local";
			this.mnuRPTVentasANuXLo.Id = 953;
			this.mnuRPTVentasANuXLo.Name = "mnuRPTVentasANuXLo";
			this.mnuRPTVentasANuXLo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasMenXLo
			// 
			this.mnuRPTVentasMenXLo.Caption = "Ventas Mensuales Por Local";
			this.mnuRPTVentasMenXLo.Id = 954;
			this.mnuRPTVentasMenXLo.Name = "mnuRPTVentasMenXLo";
			this.mnuRPTVentasMenXLo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasTarCredi
			// 
			this.mnuRPTVentasTarCredi.Caption = "Ventas Tarjetas De Credito";
			this.mnuRPTVentasTarCredi.Id = 955;
			this.mnuRPTVentasTarCredi.Name = "mnuRPTVentasTarCredi";
			this.mnuRPTVentasTarCredi.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasDetTc
			// 
			this.mnuRPTVentasDetTc.Caption = "Detalle Tarjetas De Credito";
			this.mnuRPTVentasDetTc.Id = 956;
			this.mnuRPTVentasDetTc.Name = "mnuRPTVentasDetTc";
			this.mnuRPTVentasDetTc.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasEncuestas
			// 
			this.mnuRPTVentasEncuestas.Caption = "Encuestas";
			this.mnuRPTVentasEncuestas.Id = 957;
			this.mnuRPTVentasEncuestas.Name = "mnuRPTVentasEncuestas";
			this.mnuRPTVentasEncuestas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasReCXCo
			// 
			this.mnuRPTVentasReCXCo.Caption = "Resumen de Cuentas Por Cobrar";
			this.mnuRPTVentasReCXCo.Id = 958;
			this.mnuRPTVentasReCXCo.Name = "mnuRPTVentasReCXCo";
			this.mnuRPTVentasReCXCo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasXano
			// 
			this.mnuRPTVentasXano.Caption = "Vendedor Por Año";
			this.mnuRPTVentasXano.Id = 959;
			this.mnuRPTVentasXano.Name = "mnuRPTVentasXano";
			this.mnuRPTVentasXano.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasXArtic
			// 
			this.mnuRPTVentasXArtic.Caption = "Ventas Por Articulos";
			this.mnuRPTVentasXArtic.Id = 960;
			this.mnuRPTVentasXArtic.Name = "mnuRPTVentasXArtic";
			this.mnuRPTVentasXArtic.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasArtXMes
			// 
			this.mnuRPTVentasArtXMes.Caption = "Ventas Articulos Por Mes";
			this.mnuRPTVentasArtXMes.Id = 961;
			this.mnuRPTVentasArtXMes.Name = "mnuRPTVentasArtXMes";
			this.mnuRPTVentasArtXMes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuRPTVentasArtXMes.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPTVentasArtXMes_ItemClick);
			// 
			// mnuRPTVentasPorcMargUtilidad
			// 
			this.mnuRPTVentasPorcMargUtilidad.Caption = "Porcentaje Margen Utilidad";
			this.mnuRPTVentasPorcMargUtilidad.Id = 939;
			this.mnuRPTVentasPorcMargUtilidad.Name = "mnuRPTVentasPorcMargUtilidad";
			this.mnuRPTVentasPorcMargUtilidad.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPTVentasPorcMargUtilidad_ItemClick);
			// 
			// barSubItem12
			// 
			this.barSubItem12.Caption = "Comisiones";
			this.barSubItem12.Id = 963;
			this.barSubItem12.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasComisionesConf),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasComisionesCali),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasComisionesDetGe),
																																															new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPTVentasComisionesConsulta)});
			this.barSubItem12.Name = "barSubItem12";
			this.barSubItem12.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// mnuRPTVentasComisionesConf
			// 
			this.mnuRPTVentasComisionesConf.Caption = "Configuración";
			this.mnuRPTVentasComisionesConf.Id = 964;
			this.mnuRPTVentasComisionesConf.Name = "mnuRPTVentasComisionesConf";
			// 
			// mnuRPTVentasComisionesCali
			// 
			this.mnuRPTVentasComisionesCali.Caption = "Calificación";
			this.mnuRPTVentasComisionesCali.Id = 965;
			this.mnuRPTVentasComisionesCali.Name = "mnuRPTVentasComisionesCali";
			// 
			// mnuRPTVentasComisionesDetGe
			// 
			this.mnuRPTVentasComisionesDetGe.Caption = "Detalle General";
			this.mnuRPTVentasComisionesDetGe.Id = 966;
			this.mnuRPTVentasComisionesDetGe.Name = "mnuRPTVentasComisionesDetGe";
			// 
			// mnuRPTVentasComisionesConsulta
			// 
			this.mnuRPTVentasComisionesConsulta.Caption = "Consulta";
			this.mnuRPTVentasComisionesConsulta.Id = 967;
			this.mnuRPTVentasComisionesConsulta.Name = "mnuRPTVentasComisionesConsulta";
			// 
			// mnuReporVentasPDGA
			// 
			this.mnuReporVentasPDGA.Caption = "Porcentaje De Gastos Anuales";
			this.mnuReporVentasPDGA.Id = 971;
			this.mnuReporVentasPDGA.Name = "mnuReporVentasPDGA";
			this.mnuReporVentasPDGA.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuReporVentasPDGA_ItemClick);
			// 
			// mnuReporVentasDFI
			// 
			this.mnuReporVentasDFI.Caption = "Seguro DFI";
			this.mnuReporVentasDFI.Id = 1001;
			this.mnuReporVentasDFI.Name = "mnuReporVentasDFI";
			this.mnuReporVentasDFI.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuReporVentasDFI_ItemClick);
			// 
			// mnuRPtCREPOINT
			// 
			this.mnuRPtCREPOINT.Caption = "CrediPOINT";
			this.mnuRPtCREPOINT.Id = 909;
			this.mnuRPtCREPOINT.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtCREPOINTFDG),
																																																new DevExpress.XtraBars.LinkPersistInfo(this.mnuRPtCREPOINRMD)});
			this.mnuRPtCREPOINT.Name = "mnuRPtCREPOINT";
			// 
			// mnuRPtCREPOINTFDG
			// 
			this.mnuRPtCREPOINTFDG.Caption = "Fondo De Garantía";
			this.mnuRPtCREPOINTFDG.Id = 998;
			this.mnuRPtCREPOINTFDG.Name = "mnuRPtCREPOINTFDG";
			this.mnuRPtCREPOINTFDG.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtCREPOINTFDG_ItemClick);
			// 
			// mnuRPtCREPOINRMD
			// 
			this.mnuRPtCREPOINRMD.Caption = "Rango Morosidad";
			this.mnuRPtCREPOINRMD.Id = 999;
			this.mnuRPtCREPOINRMD.Name = "mnuRPtCREPOINRMD";
			this.mnuRPtCREPOINRMD.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuRPtCREPOINRMD_ItemClick);
			// 
			// mnuRPtDT
			// 
			this.mnuRPtDT.Caption = "DT";
			this.mnuRPtDT.Id = 910;
			this.mnuRPtDT.Name = "mnuRPtDT";
			// 
			// mnuRPtPeonyServicios
			// 
			this.mnuRPtPeonyServicios.Caption = "Peony Servicios";
			this.mnuRPtPeonyServicios.Id = 911;
			this.mnuRPtPeonyServicios.Name = "mnuRPtPeonyServicios";
			// 
			// mnuIngreso
			// 
			this.mnuIngreso.Caption = "Ingreso";
			this.mnuIngreso.Id = 938;
			this.mnuIngreso.Name = "mnuIngreso";
			this.mnuIngreso.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuIngreso_ItemClick);
			// 
			// barButtonItem86
			// 
			this.barButtonItem86.Caption = "calendario";
			this.barButtonItem86.Id = 994;
			this.barButtonItem86.Name = "barButtonItem86";
			this.barButtonItem86.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.barButtonItem86.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem86_ItemClick);
			// 
			// barButtonItem87
			// 
			this.barButtonItem87.Caption = "gex";
			this.barButtonItem87.Id = 995;
			this.barButtonItem87.Name = "barButtonItem87";
			this.barButtonItem87.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem87_ItemClick);
			// 
			// bar2
			// 
			this.bar2.BarName = "Custom 2";
			this.bar2.DockCol = 0;
			this.bar2.DockRow = 1;
			this.bar2.DockStyle = DevExpress.XtraBars.BarDockStyle.Top;
			this.bar2.FloatLocation = new System.Drawing.Point(70, 189);
			this.bar2.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																											new DevExpress.XtraBars.LinkPersistInfo(this.barNuevo),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.barAbrir),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.barGrabar),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.BarPlanCuentas, true),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.barAsiento),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.BarRepContabilidad),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.BarCompras, true),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.BarVentas),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.BarArticulo),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.BarCliente),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.BarRepResumen, true),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.BarRepInventario),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.BarRepArticulo),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.barSeguridad),
																																											new DevExpress.XtraBars.LinkPersistInfo(this.barPagar),
																																											new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.barCobro, DevExpress.XtraBars.BarItemPaintStyle.Standard),
																																											new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.Width, this.lblMensaje, "", false, true, true, 85),
																																											new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.Width, this.barSkin, "", false, true, true, 103)});
			this.bar2.Text = "Barra";
			// 
			// barNuevo
			// 
			this.barNuevo.Caption = "Nuevo";
			this.barNuevo.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.barNuevo.Description = "Crea nueva Empresa";
			this.barNuevo.Hint = "Nueva Empresa";
			this.barNuevo.Id = 4;
			this.barNuevo.ImageIndex = 8;
			this.barNuevo.Name = "barNuevo";
			this.barNuevo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchNuevo_ItemClick);
			// 
			// barAbrir
			// 
			this.barAbrir.Caption = "Abrir";
			this.barAbrir.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.barAbrir.Hint = "Abre una Empresa";
			this.barAbrir.Id = 6;
			this.barAbrir.ImageIndex = 4;
			this.barAbrir.Name = "barAbrir";
			this.barAbrir.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchIniSes_ItemClick);
			// 
			// barGrabar
			// 
			this.barGrabar.Caption = "Grabar";
			this.barGrabar.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.barGrabar.Hint = "Graba respaldo de base de datos";
			this.barGrabar.Id = 218;
			this.barGrabar.ImageIndex = 9;
			this.barGrabar.Name = "barGrabar";
			this.barGrabar.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuArchGuardar_ItemClick);
			// 
			// BarPlanCuentas
			// 
			this.BarPlanCuentas.Caption = "Plan de Cuentas";
			this.BarPlanCuentas.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.BarPlanCuentas.Hint = "Plan de Cuentas";
			this.BarPlanCuentas.Id = 234;
			this.BarPlanCuentas.ImageIndex = 14;
			this.BarPlanCuentas.Name = "BarPlanCuentas";
			this.BarPlanCuentas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCont_Plan_Ctas_ItemClick);
			// 
			// barAsiento
			// 
			this.barAsiento.Caption = "Asientos";
			this.barAsiento.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.barAsiento.Hint = "Asientos de diario";
			this.barAsiento.Id = 224;
			this.barAsiento.ImageIndex = 36;
			this.barAsiento.Name = "barAsiento";
			this.barAsiento.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuCont_Asientos_ItemClick);
			// 
			// BarRepContabilidad
			// 
			this.BarRepContabilidad.Caption = "Reporte de Contabilidad";
			this.BarRepContabilidad.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.BarRepContabilidad.Hint = "Reportes de Contabilidad";
			this.BarRepContabilidad.Id = 225;
			this.BarRepContabilidad.ImageIndex = 16;
			this.BarRepContabilidad.Name = "BarRepContabilidad";
			this.BarRepContabilidad.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuContRepCont_ItemClick);
			// 
			// BarCompras
			// 
			this.BarCompras.Caption = "Compra";
			this.BarCompras.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.BarCompras.Hint = "Compras";
			this.BarCompras.Id = 226;
			this.BarCompras.ImageIndex = 18;
			this.BarCompras.Name = "BarCompras";
			this.BarCompras.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactCompras_ItemClick);
			// 
			// BarVentas
			// 
			this.BarVentas.Caption = "Venta";
			this.BarVentas.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.BarVentas.Hint = "Ventas";
			this.BarVentas.Id = 232;
			this.BarVentas.ImageIndex = 22;
			this.BarVentas.Name = "BarVentas";
			this.BarVentas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactVentas_ItemClick);
			// 
			// BarArticulo
			// 
			this.BarArticulo.Caption = "Articulo";
			this.BarArticulo.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.BarArticulo.Hint = "Artículos";
			this.BarArticulo.Id = 228;
			this.BarArticulo.ImageIndex = 26;
			this.BarArticulo.Name = "BarArticulo";
			this.BarArticulo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactArtic_ItemClick);
			// 
			// BarCliente
			// 
			this.BarCliente.Caption = "Cliente";
			this.BarCliente.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.BarCliente.Hint = "Clientes";
			this.BarCliente.Id = 229;
			this.BarCliente.ImageIndex = 21;
			this.BarCliente.Name = "BarCliente";
			this.BarCliente.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactClientes_ItemClick);
			// 
			// BarRepResumen
			// 
			this.BarRepResumen.Caption = "Resumen";
			this.BarRepResumen.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.BarRepResumen.Hint = "Resúmenes de Facturación";
			this.BarRepResumen.Id = 230;
			this.BarRepResumen.ImageIndex = 24;
			this.BarRepResumen.Name = "BarRepResumen";
			this.BarRepResumen.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.BarRepResumen.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactRepResumenes_ItemClick);
			// 
			// BarRepInventario
			// 
			this.BarRepInventario.Caption = "Inventario";
			this.BarRepInventario.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.BarRepInventario.Hint = "Reportes de Inventario";
			this.BarRepInventario.Id = 231;
			this.BarRepInventario.ImageIndex = 23;
			this.BarRepInventario.Name = "BarRepInventario";
			this.BarRepInventario.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.BarRepInventario.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactRepInventario_ItemClick);
			// 
			// BarRepArticulo
			// 
			this.BarRepArticulo.Caption = "Rep. Articulos";
			this.BarRepArticulo.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.BarRepArticulo.Hint = "Reportes de Artículos";
			this.BarRepArticulo.Id = 233;
			this.BarRepArticulo.ImageIndex = 25;
			this.BarRepArticulo.Name = "BarRepArticulo";
			this.BarRepArticulo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.BarRepArticulo.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactGraficos_ItemClick);
			// 
			// barSeguridad
			// 
			this.barSeguridad.Caption = "Seguridad";
			this.barSeguridad.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.barSeguridad.Hint = "Seguridades del sistema";
			this.barSeguridad.Id = 235;
			this.barSeguridad.ImageIndex = 13;
			this.barSeguridad.Name = "barSeguridad";
			this.barSeguridad.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.Mnu_EdicSeguridad_ItemClick);
			// 
			// barPagar
			// 
			this.barPagar.Caption = "Pagar";
			this.barPagar.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.barPagar.Description = "Cuentas por Pagar";
			this.barPagar.Hint = "Cuentas por Pagar";
			this.barPagar.Id = 239;
			this.barPagar.ImageIndex = 45;
			this.barPagar.Name = "barPagar";
			this.barPagar.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactCtasPagar_ItemClick);
			// 
			// barCobro
			// 
			this.barCobro.Caption = "Cobros";
			this.barCobro.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.barCobro.Description = "Cuentas por Cobrar";
			this.barCobro.Hint = "Cuentas por Cobrar";
			this.barCobro.Id = 238;
			this.barCobro.ImageIndex = 47;
			this.barCobro.Name = "barCobro";
			this.barCobro.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactCtasCobrar_ItemClick);
			// 
			// lblMensaje
			// 
			this.lblMensaje.CanOpenEdit = false;
			this.lblMensaje.Caption = "barEditItem1";
			this.lblMensaje.Edit = this.repositoryItemTextEdit1;
			this.lblMensaje.EditValue = "Tiempo: 0.00";
			this.lblMensaje.Id = 437;
			this.lblMensaje.Name = "lblMensaje";
			this.lblMensaje.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// repositoryItemTextEdit1
			// 
			this.repositoryItemTextEdit1.AutoHeight = false;
			this.repositoryItemTextEdit1.Name = "repositoryItemTextEdit1";
			// 
			// barSkin
			// 
			this.barSkin.Edit = this.repositoryItemComboBox1;
			this.barSkin.EditValue = "Rotativo";
			this.barSkin.Id = 438;
			this.barSkin.Name = "barSkin";
			this.barSkin.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.barSkin.EditValueChanged += new System.EventHandler(this.barEditItem1_EditValueChanged);
			// 
			// repositoryItemComboBox1
			// 
			this.repositoryItemComboBox1.AutoHeight = false;
			this.repositoryItemComboBox1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemComboBox1.Items.AddRange(new object[] {
																																 "Rotativo",
																																 "Caramel",
																																 "Money Twins",
																																 "Lilian",
																																 "Blue",
																																 "Imaginary",
																																 "The Asphalt World",
																																 "Black"});
			this.repositoryItemComboBox1.Name = "repositoryItemComboBox1";
			// 
			// MnuHerr_Cre_Sal_Ini
			// 
			this.MnuHerr_Cre_Sal_Ini.Caption = "Crea Saldo Inicial";
			this.MnuHerr_Cre_Sal_Ini.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuHerr_Cre_Sal_Ini.Id = 68;
			this.MnuHerr_Cre_Sal_Ini.Name = "MnuHerr_Cre_Sal_Ini";
			// 
			// MnuHerr_Cop_Sal_Ini
			// 
			this.MnuHerr_Cop_Sal_Ini.Caption = "Copia Saldo Inicial";
			this.MnuHerr_Cop_Sal_Ini.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.MnuHerr_Cop_Sal_Ini.Id = 69;
			this.MnuHerr_Cop_Sal_Ini.Name = "MnuHerr_Cop_Sal_Ini";
			// 
			// MnuContRepGov
			// 
			this.MnuContRepGov.Caption = "Reportes NIIF";
			this.MnuContRepGov.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContRepGov.Id = 286;
			this.MnuContRepGov.Name = "MnuContRepGov";
			this.MnuContRepGov.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// MnuContFormulario
			// 
			this.MnuContFormulario.Caption = "Formularios";
			this.MnuContFormulario.CategoryGuid = new System.Guid("a520686f-b7d6-4ade-a6e1-d177f34cfea3");
			this.MnuContFormulario.Id = 101;
			this.MnuContFormulario.ItemShortcut = new DevExpress.XtraBars.BarShortcut((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F));
			this.MnuContFormulario.Name = "MnuContFormulario";
			this.MnuContFormulario.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// MnuFactPedEmpleados
			// 
			this.MnuFactPedEmpleados.Caption = "Pedidos Empleados";
			this.MnuFactPedEmpleados.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.MnuFactPedEmpleados.Id = 115;
			this.MnuFactPedEmpleados.Name = "MnuFactPedEmpleados";
			this.MnuFactPedEmpleados.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactPedEmpleados_ItemClick);
			// 
			// mnuFactLotOrdenMov
			// 
			this.mnuFactLotOrdenMov.Caption = "Orden Movilizacion";
			this.mnuFactLotOrdenMov.CategoryGuid = new System.Guid("4a0dee81-61cf-46d2-9711-1c51694350b9");
			this.mnuFactLotOrdenMov.Id = 251;
			this.mnuFactLotOrdenMov.Name = "mnuFactLotOrdenMov";
			// 
			// MnuRolDatGen
			// 
			this.MnuRolDatGen.Caption = "Anticipos";
			this.MnuRolDatGen.CategoryGuid = new System.Guid("dab6fe3e-8b7a-4748-9620-cf654702690b");
			this.MnuRolDatGen.Id = 141;
			this.MnuRolDatGen.Name = "MnuRolDatGen";
			this.MnuRolDatGen.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// MnuProducPrecios
			// 
			this.MnuProducPrecios.Caption = "Precios";
			this.MnuProducPrecios.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.MnuProducPrecios.Id = 153;
			this.MnuProducPrecios.Name = "MnuProducPrecios";
			// 
			// MnuAyuda
			// 
			this.MnuAyuda.Caption = "A&yuda";
			this.MnuAyuda.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.MnuAyuda.Id = 178;
			this.MnuAyuda.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																													new DevExpress.XtraBars.LinkPersistInfo(this.MnuAyuda_Ayuda),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.MnuAyudaRegist),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.MnuAyudaAcercaCIAD, true),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.mnuAyuLiberar),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.btControlRemoto, false),
																																													new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.None, false, this.mnuAyuJuego, false),
																																													new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem63)});
			this.MnuAyuda.Name = "MnuAyuda";
			// 
			// MnuAyuda_Ayuda
			// 
			this.MnuAyuda_Ayuda.Caption = "Ayuda     ";
			this.MnuAyuda_Ayuda.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.MnuAyuda_Ayuda.Id = 195;
			this.MnuAyuda_Ayuda.ItemShortcut = new DevExpress.XtraBars.BarShortcut(System.Windows.Forms.Keys.F1);
			this.MnuAyuda_Ayuda.Name = "MnuAyuda_Ayuda";
			this.MnuAyuda_Ayuda.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.MnuAyuda_Ayuda.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuAyuda_Ayuda_ItemClick);
			// 
			// MnuAyudaRegist
			// 
			this.MnuAyudaRegist.Caption = "Registro";
			this.MnuAyudaRegist.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.MnuAyudaRegist.Id = 199;
			this.MnuAyudaRegist.Name = "MnuAyudaRegist";
			this.MnuAyudaRegist.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.MnuAyudaRegist.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuAyudaRegist_ItemClick);
			// 
			// MnuAyudaAcercaCIAD
			// 
			this.MnuAyudaAcercaCIAD.Caption = "Acerca de Latinium";
			this.MnuAyudaAcercaCIAD.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.MnuAyudaAcercaCIAD.Id = 200;
			this.MnuAyudaAcercaCIAD.Name = "MnuAyudaAcercaCIAD";
			this.MnuAyudaAcercaCIAD.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.MnuAyudaAcercaCIAD.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuAyudaAcercaCIAD_ItemClick);
			// 
			// mnuAyuLiberar
			// 
			this.mnuAyuLiberar.Caption = "Liberar Memoria";
			this.mnuAyuLiberar.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.mnuAyuLiberar.Id = 277;
			this.mnuAyuLiberar.Name = "mnuAyuLiberar";
			this.mnuAyuLiberar.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuAyuLiberar.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuAyuLiberar_ItemClick);
			// 
			// btControlRemoto
			// 
			this.btControlRemoto.Caption = "Ayuda Remota";
			this.btControlRemoto.Id = 417;
			this.btControlRemoto.Name = "btControlRemoto";
			this.btControlRemoto.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.btControlRemoto.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btControlRemoto_ItemClick);
			// 
			// mnuAyuJuego
			// 
			this.mnuAyuJuego.Caption = "Juego";
			this.mnuAyuJuego.Id = 452;
			this.mnuAyuJuego.Name = "mnuAyuJuego";
			this.mnuAyuJuego.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuAyuJuego.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuAyuJuego_ItemClick);
			// 
			// barButtonItem63
			// 
			this.barButtonItem63.Caption = "Politicas de Credito y Garantias";
			this.barButtonItem63.Id = 818;
			this.barButtonItem63.Name = "barButtonItem63";
			// 
			// mnuHerrSegReqEst
			// 
			this.mnuHerrSegReqEst.Caption = "Requerimiento Estado";
			this.mnuHerrSegReqEst.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuHerrSegReqEst.Id = 398;
			this.mnuHerrSegReqEst.Name = "mnuHerrSegReqEst";
			this.mnuHerrSegReqEst.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrSegReqEst_ItemClick);
			// 
			// mnuProdOrdenT
			// 
			this.mnuProdOrdenT.Caption = "Orden de Trabajo";
			this.mnuProdOrdenT.CategoryGuid = new System.Guid("4889142b-ad08-4f8e-ba2d-15b0fc1cb0e9");
			this.mnuProdOrdenT.Id = 275;
			this.mnuProdOrdenT.Name = "mnuProdOrdenT";
			// 
			// mneArchImpFidelius
			// 
			this.mneArchImpFidelius.Caption = "Micros Fidelius";
			this.mneArchImpFidelius.CategoryGuid = new System.Guid("d5bcd0b6-2b08-41d2-8022-e8963b4c4759");
			this.mneArchImpFidelius.Id = 293;
			this.mneArchImpFidelius.Name = "mneArchImpFidelius";
			this.mneArchImpFidelius.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mneArchImpFidelius_ItemClick);
			// 
			// mnuHerrSegTransf
			// 
			this.mnuHerrSegTransf.Caption = "Transformacion";
			this.mnuHerrSegTransf.CategoryGuid = new System.Guid("7839a577-2a4b-4cd0-abc9-380ef100654c");
			this.mnuHerrSegTransf.Id = 309;
			this.mnuHerrSegTransf.Name = "mnuHerrSegTransf";
			this.mnuHerrSegTransf.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mbuHerrSegTransf_ItemClick);
			// 
			// mnuHerrContratosAdemdum
			// 
			this.mnuHerrContratosAdemdum.Caption = "Ademdums";
			this.mnuHerrContratosAdemdum.Id = 411;
			this.mnuHerrContratosAdemdum.Name = "mnuHerrContratosAdemdum";
			// 
			// mnuArchImpMarceloVentas
			// 
			this.mnuArchImpMarceloVentas.Caption = "Marcelo General Ventas";
			this.mnuArchImpMarceloVentas.Id = 424;
			this.mnuArchImpMarceloVentas.Name = "mnuArchImpMarceloVentas";
			this.mnuArchImpMarceloVentas.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btMarceloGeneralVentas_ItemClick);
			// 
			// MnuFactLotimp2
			// 
			this.MnuFactLotimp2.Id = 439;
			this.MnuFactLotimp2.Name = "MnuFactLotimp2";
			// 
			// MnuFactContometros
			// 
			this.MnuFactContometros.Caption = "Contometros";
			this.MnuFactContometros.Id = 443;
			this.MnuFactContometros.Name = "MnuFactContometros";
			this.MnuFactContometros.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFactContometros_ItemClick);
			// 
			// barButtonItem3
			// 
			this.barButtonItem3.Caption = "Gruas";
			this.barButtonItem3.Id = 444;
			this.barButtonItem3.Name = "barButtonItem3";
			// 
			// barButtonItem4
			// 
			this.barButtonItem4.Caption = "barButtonItem4";
			this.barButtonItem4.Id = 447;
			this.barButtonItem4.Name = "barButtonItem4";
			// 
			// mnuHerrAgenda
			// 
			this.mnuHerrAgenda.Caption = "Alarmas";
			this.mnuHerrAgenda.Id = 449;
			this.mnuHerrAgenda.Name = "mnuHerrAgenda";
			this.mnuHerrAgenda.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuHerrAgenda_ItemClick);
			// 
			// MnuDepTec
			// 
			this.MnuDepTec.Caption = "&Dep. Tecnico";
			this.MnuDepTec.Id = 459;
			this.MnuDepTec.Name = "MnuDepTec";
			// 
			// barButtonItem5
			// 
			this.barButtonItem5.Id = 465;
			this.barButtonItem5.Name = "barButtonItem5";
			// 
			// MnuHerrAlarma
			// 
			this.MnuHerrAlarma.Caption = "Alarma";
			this.MnuHerrAlarma.Id = 463;
			this.MnuHerrAlarma.Name = "MnuHerrAlarma";
			this.MnuHerrAlarma.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.MnuHerrAlarma.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuHerrAlarma_ItemClick);
			// 
			// MnuFacSpif
			// 
			this.MnuFacSpif.Caption = "Spif";
			this.MnuFacSpif.Id = 464;
			this.MnuFacSpif.Name = "MnuFacSpif";
			this.MnuFacSpif.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.MnuFacSpif_ItemClick);
			// 
			// barButtonItem6
			// 
			this.barButtonItem6.Caption = "Credito";
			this.barButtonItem6.Id = 469;
			this.barButtonItem6.Name = "barButtonItem6";
			this.barButtonItem6.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem6_ItemClick);
			// 
			// barButtonItem11
			// 
			this.barButtonItem11.Caption = "Catálogos";
			this.barButtonItem11.Id = 479;
			this.barButtonItem11.Name = "barButtonItem11";
			// 
			// MnuFacOCImp
			// 
			this.MnuFacOCImp.Caption = "Orden de Compra Importacion";
			this.MnuFacOCImp.Id = 485;
			this.MnuFacOCImp.Name = "MnuFacOCImp";
			// 
			// barSubItem6
			// 
			this.barSubItem6.Caption = "C&ompras";
			this.barSubItem6.Id = 501;
			this.barSubItem6.Name = "barSubItem6";
			// 
			// barSubItem7
			// 
			this.barSubItem7.Caption = "&Ventas";
			this.barSubItem7.Id = 502;
			this.barSubItem7.Name = "barSubItem7";
			// 
			// FacGuiaRemision
			// 
			this.FacGuiaRemision.Caption = "Guia de Remisión";
			this.FacGuiaRemision.Id = 508;
			this.FacGuiaRemision.Name = "FacGuiaRemision";
			this.FacGuiaRemision.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.FacGuiaRemision_ItemClick);
			// 
			// barButtonItem16
			// 
			this.barButtonItem16.Caption = "C&ompras";
			this.barButtonItem16.Id = 524;
			this.barButtonItem16.Name = "barButtonItem16";
			// 
			// mnuVentasComDetallePersonal
			// 
			this.mnuVentasComDetallePersonal.Caption = "Asignaciones";
			this.mnuVentasComDetallePersonal.Id = 545;
			this.mnuVentasComDetallePersonal.Name = "mnuVentasComDetallePersonal";
			this.mnuVentasComDetallePersonal.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasComDetallePersonal_ItemClick);
			// 
			// barButtonItem7
			// 
			this.barButtonItem7.Caption = "Prueba Ventas";
			this.barButtonItem7.Id = 549;
			this.barButtonItem7.Name = "barButtonItem7";
			// 
			// barButtonItem8
			// 
			this.barButtonItem8.Caption = "Prueba de Compras";
			this.barButtonItem8.Id = 551;
			this.barButtonItem8.Name = "barButtonItem8";
			// 
			// barButtonItem10
			// 
			this.barButtonItem10.Caption = "Reportes";
			this.barButtonItem10.Id = 557;
			this.barButtonItem10.Name = "barButtonItem10";
			// 
			// mnuBodegaReportes
			// 
			this.mnuBodegaReportes.ButtonStyle = DevExpress.XtraBars.BarButtonStyle.DropDown;
			this.mnuBodegaReportes.Caption = "Reportes";
			this.mnuBodegaReportes.Id = 567;
			this.mnuBodegaReportes.Name = "mnuBodegaReportes";
			// 
			// mnuTomaInventario
			// 
			this.mnuTomaInventario.Caption = "Toma de Inventario";
			this.mnuTomaInventario.Id = 568;
			this.mnuTomaInventario.Name = "mnuTomaInventario";
			// 
			// barButtonItem13
			// 
			this.barButtonItem13.Caption = "Bancos";
			this.barButtonItem13.Id = 574;
			this.barButtonItem13.Name = "barButtonItem13";
			// 
			// barButtonItem9
			// 
			this.barButtonItem9.Caption = "Transferencia Multiple (Pruebas)";
			this.barButtonItem9.Id = 585;
			this.barButtonItem9.Name = "barButtonItem9";
			// 
			// mnuCmpRepVM
			// 
			this.mnuCmpRepVM.Caption = "Ventas Por Marca";
			this.mnuCmpRepVM.Id = 588;
			this.mnuCmpRepVM.Name = "mnuCmpRepVM";
			// 
			// barButtonItem19
			// 
			this.barButtonItem19.Caption = "barButtonItem19";
			this.barButtonItem19.Id = 595;
			this.barButtonItem19.Name = "barButtonItem19";
			// 
			// barButtonItem20
			// 
			this.barButtonItem20.Caption = "P L P";
			this.barButtonItem20.Id = 621;
			this.barButtonItem20.Name = "barButtonItem20";
			this.barButtonItem20.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			// 
			// barToolbarsListItem1
			// 
			this.barToolbarsListItem1.Caption = "aaa";
			this.barToolbarsListItem1.Id = 623;
			this.barToolbarsListItem1.Name = "barToolbarsListItem1";
			// 
			// barButtonItem21
			// 
			this.barButtonItem21.Caption = "Habitaciones";
			this.barButtonItem21.Id = 624;
			this.barButtonItem21.Name = "barButtonItem21";
			// 
			// barSubItem8
			// 
			this.barSubItem8.Caption = "aaa";
			this.barSubItem8.Id = 625;
			this.barSubItem8.Name = "barSubItem8";
			// 
			// barButtonItem23
			// 
			this.barButtonItem23.Caption = "Reportes";
			this.barButtonItem23.Id = 652;
			this.barButtonItem23.Name = "barButtonItem23";
			// 
			// mnuVentasVentasCP
			// 
			this.mnuVentasVentasCP.Caption = "Cooperativa";
			this.mnuVentasVentasCP.Id = 657;
			this.mnuVentasVentasCP.Name = "mnuVentasVentasCP";
			this.mnuVentasVentasCP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			this.mnuVentasVentasCP.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.mnuVentasVentasCP_ItemClick);
			// 
			// barButtonItem33
			// 
			this.barButtonItem33.Caption = "Nomina Pruebas";
			this.barButtonItem33.Id = 677;
			this.barButtonItem33.Name = "barButtonItem33";
			// 
			// barButtonItem34
			// 
			this.barButtonItem34.Caption = "Presupuesto Por Locales";
			this.barButtonItem34.Id = 678;
			this.barButtonItem34.Name = "barButtonItem34";
			// 
			// barButtonItem38
			// 
			this.barButtonItem38.Caption = "Reporte de Transferencias";
			this.barButtonItem38.Id = 685;
			this.barButtonItem38.Name = "barButtonItem38";
			// 
			// barButtonItem39
			// 
			this.barButtonItem39.Caption = "Cobranza";
			this.barButtonItem39.Id = 695;
			this.barButtonItem39.Name = "barButtonItem39";
			// 
			// barButtonItem41
			// 
			this.barButtonItem41.Caption = "Reportes";
			this.barButtonItem41.Id = 707;
			this.barButtonItem41.Name = "barButtonItem41";
			// 
			// barButtonItem44
			// 
			this.barButtonItem44.Caption = "Reportes";
			this.barButtonItem44.Id = 721;
			this.barButtonItem44.Name = "barButtonItem44";
			// 
			// barButtonItem45
			// 
			this.barButtonItem45.Caption = "Lotes";
			this.barButtonItem45.Id = 723;
			this.barButtonItem45.Name = "barButtonItem45";
			// 
			// barButtonItem46
			// 
			this.barButtonItem46.Caption = "Pichincha";
			this.barButtonItem46.Id = 725;
			this.barButtonItem46.Name = "barButtonItem46";
			// 
			// barButtonItem47
			// 
			this.barButtonItem47.Caption = "Solidario";
			this.barButtonItem47.Id = 726;
			this.barButtonItem47.Name = "barButtonItem47";
			// 
			// barButtonItem48
			// 
			this.barButtonItem48.Caption = "Pichincha";
			this.barButtonItem48.Id = 728;
			this.barButtonItem48.Name = "barButtonItem48";
			// 
			// barButtonItem49
			// 
			this.barButtonItem49.Caption = "barButtonItem49";
			this.barButtonItem49.Id = 731;
			this.barButtonItem49.Name = "barButtonItem49";
			// 
			// barButtonItem55
			// 
			this.barButtonItem55.Caption = "barButtonItem55";
			this.barButtonItem55.Id = 741;
			this.barButtonItem55.Name = "barButtonItem55";
			// 
			// barButtonItem56
			// 
			this.barButtonItem56.Caption = "barButtonItem56";
			this.barButtonItem56.Id = 746;
			this.barButtonItem56.Name = "barButtonItem56";
			// 
			// mnuDTIRPPL
			// 
			this.mnuDTIRPPL.Caption = "Productos Pendientes en LOCALES";
			this.mnuDTIRPPL.Id = 753;
			this.mnuDTIRPPL.Name = "mnuDTIRPPL";
			// 
			// barButtonItem61
			// 
			this.barButtonItem61.Caption = "barButtonItem61";
			this.barButtonItem61.Id = 756;
			this.barButtonItem61.Name = "barButtonItem61";
			// 
			// barButtonItem70
			// 
			this.barButtonItem70.Caption = "barButtonItem70";
			this.barButtonItem70.Id = 771;
			this.barButtonItem70.Name = "barButtonItem70";
			// 
			// barButtonItem71
			// 
			this.barButtonItem71.Caption = "barButtonItem71";
			this.barButtonItem71.Id = 772;
			this.barButtonItem71.Name = "barButtonItem71";
			// 
			// mnuVRE
			// 
			this.mnuVRE.Caption = "Encuestas";
			this.mnuVRE.Id = 778;
			this.mnuVRE.Name = "mnuVRE";
			// 
			// barButtonItem74
			// 
			this.barButtonItem74.Caption = "Procesos";
			this.barButtonItem74.Id = 781;
			this.barButtonItem74.Name = "barButtonItem74";
			// 
			// barButtonItem58
			// 
			this.barButtonItem58.Caption = "Metricas";
			this.barButtonItem58.Id = 811;
			this.barButtonItem58.Name = "barButtonItem58";
			// 
			// barButtonItem59
			// 
			this.barButtonItem59.Caption = "RRHH";
			this.barButtonItem59.Id = 813;
			this.barButtonItem59.Name = "barButtonItem59";
			// 
			// barButtonItem52
			// 
			this.barButtonItem52.Caption = "barButtonItem52";
			this.barButtonItem52.Id = 851;
			this.barButtonItem52.Name = "barButtonItem52";
			// 
			// mnuBodegaTIM
			// 
			this.mnuBodegaTIM.Caption = "Turnos Ingreso de Mercaderia";
			this.mnuBodegaTIM.Id = 855;
			this.mnuBodegaTIM.Name = "mnuBodegaTIM";
			// 
			// mnuRolesComisionesPCS
			// 
			this.mnuRolesComisionesPCS.Caption = "Porcentajes De Comisiones Semanales";
			this.mnuRolesComisionesPCS.Id = 859;
			this.mnuRolesComisionesPCS.Name = "mnuRolesComisionesPCS";
			// 
			// barButtonItem43
			// 
			this.barButtonItem43.Caption = "barButtonItem43";
			this.barButtonItem43.Id = 884;
			this.barButtonItem43.Name = "barButtonItem43";
			// 
			// barButtonItem54
			// 
			this.barButtonItem54.Caption = "barButtonItem54";
			this.barButtonItem54.Id = 888;
			this.barButtonItem54.Name = "barButtonItem54";
			// 
			// barButtonItem64
			// 
			this.barButtonItem64.Caption = "Reporteria";
			this.barButtonItem64.Id = 896;
			this.barButtonItem64.Name = "barButtonItem64";
			// 
			// barLinkContainerItem1
			// 
			this.barLinkContainerItem1.Caption = "Reporteria";
			this.barLinkContainerItem1.Id = 897;
			this.barLinkContainerItem1.Name = "barLinkContainerItem1";
			// 
			// barButtonItem65
			// 
			this.barButtonItem65.Caption = "Bodega";
			this.barButtonItem65.Id = 899;
			this.barButtonItem65.Name = "barButtonItem65";
			// 
			// barButtonItem77
			// 
			this.barButtonItem77.Caption = "Roles";
			this.barButtonItem77.Id = 902;
			this.barButtonItem77.Name = "barButtonItem77";
			// 
			// barButtonItem78
			// 
			this.barButtonItem78.Caption = "Roles";
			this.barButtonItem78.Id = 903;
			this.barButtonItem78.Name = "barButtonItem78";
			// 
			// barButtonItem79
			// 
			this.barButtonItem79.Caption = "Compras";
			this.barButtonItem79.Id = 905;
			this.barButtonItem79.Name = "barButtonItem79";
			// 
			// barButtonItem80
			// 
			this.barButtonItem80.Caption = "barButtonItem80";
			this.barButtonItem80.Id = 907;
			this.barButtonItem80.Name = "barButtonItem80";
			// 
			// barButtonItem81
			// 
			this.barButtonItem81.Caption = "barButtonItem81";
			this.barButtonItem81.Id = 962;
			this.barButtonItem81.Name = "barButtonItem81";
			// 
			// barButtonItem83
			// 
			this.barButtonItem83.Caption = "barButtonItem83";
			this.barButtonItem83.Id = 969;
			this.barButtonItem83.Name = "barButtonItem83";
			// 
			// barButtonItem85
			// 
			this.barButtonItem85.Caption = "barButtonItem85";
			this.barButtonItem85.Id = 991;
			this.barButtonItem85.Name = "barButtonItem85";
			// 
			// barButtonItem88
			// 
			this.barButtonItem88.Id = 997;
			this.barButtonItem88.Name = "barButtonItem88";
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.FillOnRequest = false;
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// helpProvider1
			// 
			this.helpProvider1.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// timer
			// 
			this.timer.Enabled = true;
			this.timer.Interval = 10000;
			this.timer.Tick += new System.EventHandler(this.timer_Tick);
			// 
			// skinPrincipal
			// 
			this.skinPrincipal.LookAndFeel.SkinName = "Black";
			// 
			// barSubItem2
			// 
			this.barSubItem2.Caption = "A&yuda";
			this.barSubItem2.CategoryGuid = new System.Guid("cd830c60-7e0f-450d-aebd-b1d882d3dd88");
			this.barSubItem2.Id = 178;
			this.barSubItem2.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.MnuAyuda_Ayuda),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.MnuAyudaRegist),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.MnuAyudaAcercaCIAD, true),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuAyuLiberar),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.btControlRemoto),
																																														 new DevExpress.XtraBars.LinkPersistInfo(this.mnuAyuJuego)});
			this.barSubItem2.Name = "barSubItem2";
			// 
			// barButtonItem15
			// 
			this.barButtonItem15.Caption = "C&ompras";
			this.barButtonItem15.Id = 523;
			this.barButtonItem15.Name = "barButtonItem15";
			// 
			// MenuLatinium
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
			this.ClientSize = new System.Drawing.Size(880, 542);
			this.Controls.Add(this.barDockControlLeft);
			this.Controls.Add(this.barDockControlRight);
			this.Controls.Add(this.barDockControlBottom);
			this.Controls.Add(this.barDockControlTop);
			this.helpProvider1.SetHelpKeyword(this, "Abrir.Htm");
			this.helpProvider1.SetHelpNavigator(this, System.Windows.Forms.HelpNavigator.Index);
			this.helpProvider1.SetHelpString(this, "");
			this.IsMdiContainer = true;
			this.KeyPreview = true;
			this.Name = "MenuLatinium";
			this.helpProvider1.SetShowHelp(this, true);
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.MenuLatinium_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.MenuLatinium_Closing);
			this.Load += new System.EventHandler(this.Form1_Load);
			((System.ComponentModel.ISupportInitialize)(this.lblTiempos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemTextEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemComboBox1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void MnuCont_Bcos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Bancos miBanco=new Bancos();
			miBanco.MdiParent=this;
			miBanco.Show();
			Cursor = Cursors.Default;
		}

		private void MnuCont_Asig_Ctas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			AsignaCuentaArbol miAsignaCuenta=new AsignaCuentaArbol();
			miAsignaCuenta.MdiParent=this;
			miAsignaCuenta.Show();
			Cursor = Cursors.Default;
		}

		private void MnuCont_Ctas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			AsientoCuenta miCuentas=new AsientoCuenta();
			miCuentas.MdiParent=this;
			miCuentas.Show();
			Cursor = Cursors.Default;
		}

		private void MnuCont_Bloq_Asi_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			AsientoBloqueo miBloqueaAs = new AsientoBloqueo();
			miBloqueaAs.MdiParent = this;
			miBloqueaAs.Show();
			Cursor = Cursors.Default;
		}

		private void MnuCont_Asientos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAsientos Asientos = new frmAsientos();
			Asientos.MdiParent = this;
			Asientos.Show();
		}

		private void MnuCont_Cier_Per_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmPeriodos Periodo = new frmPeriodos();
			Periodo.MdiParent = this;
			Periodo.Show();
			Cursor = Cursors.Default;
		}

		private void MnuArch_Impo_CIAD_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ImportarLatinium miLatinium = new ImportarLatinium();
			miLatinium.MdiParent = this;
			miLatinium.Show();
			Cursor = Cursors.Default;
		}

		private void MnuArch_Impo_Otros_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ImportarExcel miImportar=new ImportarExcel();
			miImportar.MdiParent=this;
			miImportar.Show();
			Cursor = Cursors.Default;
		}

		private void MnuArch_Resp_Comp_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Respaldo miRespaldo=new Respaldo();
			miRespaldo.MdiParent=this;
			miRespaldo.Show();
		}

		private void MnuArchProp_Cont_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			SeteoC miPropContabilidad=new SeteoC();
			miPropContabilidad.MdiParent=this;
			miPropContabilidad.Show();
			Cursor = Cursors.Default;
		}

		private void MnuArchPropFact_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			SeteoF miProp_Facturación=new SeteoF();
			miProp_Facturación.MdiParent=this;
			miProp_Facturación.Show();
			Cursor = Cursors.Default;
		}

		
		private void MnuArchProp_Rol_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			SeteoR miRol = new SeteoR();
			miRol.MdiParent = this;
			miRol.Show();
		}

		public void AsignaMenu(C1DataTable dt, int iNivel)
		{
			bool [] a = new bool[2000];
			int iPos = 0;
			int i;
			for(i=0; i<2000; i++) a[i] = true;
			string stNivel = "n_" + iNivel.ToString();
			foreach(C1DataRow dr in dt.Rows)
			{
				if (dr["Tipo"].Equals(System.DBNull.Value)) continue;
				iPos = int.Parse(dr["Tipo"].ToString());
				if (iPos == 1 || iPos == 2)
				{
					try
					{
						iPos = int.Parse(dr["pos"].ToString());
						a[iPos] = bool.Parse(dr[stNivel].ToString());
					}
					catch (Exception ex)
					{
						MessageBox.Show(ex.Message + " idSeg " + dr["id_seg"].ToString());
					}
				}
			}
			AsignaVector(a);
			VisibleMenu(dt);
		}
			
		public void VisibleMenu(C1DataTable dt)
		{
			bool [] a = new bool[2000];
			int iPos = 0;
			int i;
			for(i=0; i<2000; i++) a[i] = true;
			string stNivel = "Vn_" + iNivel.ToString();//"Visible";
			foreach(C1DataRow dr in dt.Rows)
			{
				if (dr["Tipo"].Equals(System.DBNull.Value)) continue;
				iPos = int.Parse(dr["Tipo"].ToString());
				if (iPos == 1 || iPos == 2)
				{
					try
					{
						iPos = int.Parse(dr["pos"].ToString());
						a[iPos] = bool.Parse(dr[stNivel].ToString());
					}
					catch (Exception ex)
					{
						MessageBox.Show(ex.Message + " idSeg " + dr["id_seg"].ToString());
					}
				}
			}
			AsignaVectorV(a);
		}
		
		bool bAbrePromocionesActivas = false;
		bool bAbreListaPersonalPrueba = false;
		bool bAbreListaClientesRevisados = false;
		bool bAbreListaClientesAtrasados = false;

		/// <summary>
		/// Habilita Los Menus
		/// </summary>
		/// <param name="a"></param>
		private void AsignaVector(bool [] a)
		{
			#region Archivo
			this.MnuArch.Enabled =a[1];
			//			this.MnuArchNuevo.Enabled =a[11];
			//			this.barNuevo.Enabled =a[11];
			this.MnuArchAbrir.Enabled =a[12];
			this.barAbrir.Enabled = a[12];
			this.MnuArchGuardar.Enabled =a[13];
			this.barGrabar.Enabled =a[13];			
			this.MnuArchImpoExpo.Enabled=a[14];
			this.MnuArchImpoLat.Enabled=a[17];
			this.MnuArchImpoOtros.Enabled=a[18];
			this.MnuArchProp.Enabled=a[15];
			this.mnuArchPropGen.Enabled=a[19];
			this.MnuArchPropCont.Enabled=a[20];
			this.MnuArchPropFact.Enabled=a[21];
			this.MnuArchPropRol.Enabled=a[22];			
			this.mnuArchPropEmpresa.Enabled=a[23];
			this.mnuArchPropConfig.Enabled=a[24];
			this.MnuArchConfImp.Enabled = a[16];
			#endregion Archivo
			#region Herramientas
			this.MnuHerr_.Enabled = a[2];
			this.MnuHerr_Camb_Contr.Enabled = a[25];
			this.MnuHerr_Actu_Vers.Enabled = a[26];
			this.MnuHerr_Act_Saldo.Enabled = a[31];
			this.MnuHerr_Camb_Ctes.Enabled = a[28];
			this.MnuHerr_Copi_Arch.Enabled = a[29];
			this.MnuHerr_Auditoría.Enabled = a[30];
			bAuditoria = a[30];
			this.MnuHerr_Sal_Ini.Enabled = a[26];
			this.MnuEdicSeguridad.Enabled=a[32];
			this.MnuEdicSegMenu.Enabled=a[33];
			this.barSeguridad.Enabled=a[33];
			this.MnuEdicSegUsuario.Enabled=a[34];
			this.mnuHerramientasNumeracion.Enabled=a[433];
			this.mnuHerramientasCRIVA.Enabled=a[568];
			//this.mnuHerrReporte.Enabled = a[32];
			this.mnuHerrSegCD.Enabled=a[544];
			this.mnuHerrALF.Enabled=a[602];
			this.mnuHerramientasPTP.Enabled=a[733];
			this.mnuHerramientasAP.Enabled=a[734];
			this.mnuHerramientasPTA.Enabled=a[908];
			this.mnuHerramientasAPL.Enabled=a[928];
			this.mnuHerramientasASDP.Enabled=a[929];
			
			
			#endregion Herramientas
			#region Bodega
			this.MnuBodega.Enabled = a[3];
			this.mnuBodegaTRM.Enabled = a[881];
			this.MnuFactSerie.Enabled = a[35];
			this.MnuFacSeriales.Enabled = a[36];
			this.mnuFactBodegas.Enabled = a[42];
			this.mnuFactGuiaRemision.Enabled = a[39];
			this.mnuBodegaGRC.Enabled = a[39];
			this.MnuFactEnvTransferencia.Enabled = a[40];
			this.MnuFactRecTransferencia.Enabled = a[41];
			this.MnuFacInv.Enabled=a[37];
			this.mnuBodegaKardex.Enabled = a[38];
			this.mnuBodegaCI.Enabled=a[541];
			this.mnuBodegaDCI.Enabled = a[542];
			this.mnuBodegaCierreInventario.Enabled = a[551];
			this.mnuBodegaIF.Enabled = a[569];
			this.mnuBodegaInvVal.Enabled = a[685];
			this.mnuBodegaSectorizacion.Enabled = a[579];
			this.mnuBodegaIngSeriales.Enabled = a[603];
			this.mnuComprasTEL.Enabled = a[682];
			this.mnuReportesBodega.Enabled = a[683];
			this.mnuRepCmpTransf.Enabled = a[684];
			this.mnuBodegaIPR.Enabled = a[696];
			#endregion Bodega
			#region Contabilidad
			this.MnuCont.Enabled=a[4];
			this.MnuContPlanCtas.Enabled=a[43];
			this.BarPlanCuentas.Enabled=a[43];
			this.MnuContAsientos.Enabled=a[44];
			this.barAsiento.Enabled=a[44];
			this.MnuContAsigCtas.Enabled=a[46];
			this.mnuContReportes.Enabled=a[47];
			this.MnuContRepCont.Enabled=a[45];
			this.BarRepContabilidad.Enabled=a[45];
			this.MnuContTransacciones.Enabled=a[48];
			this.MnuContTransE.Enabled=a[73];
			this.MnuContTransI.Enabled=a[74];
			this.MnuContTransR.Enabled=a[75];
			this.MnuContSRI.Enabled=a[49];
			this.mnuContTransAnexoTr.Enabled=a[76];
			this.mnuContTransFormulario.Enabled=a[77];
			this.mnuContTransConciliacion.Enabled=a[78];
			this.MnuContCXP.Enabled=a[50];
			this.barPagar.Enabled = a[50];
			this.MnuFactAutPagos.Enabled=a[51];
			this.MnuFacPagosAut.Enabled=a[52];
			this.MnuContCXC.Enabled=a[53];
			this.barCobro.Enabled = a[53];
			this.MnuFacArtServicios.Enabled=a[54];
			this.MnuFactActFijos.Enabled=a[55];
			this.MnuContEstadoCta.Enabled=a[58];
			this.MnuContCtas.Enabled=a[57];
			this.MnuContBcos.Enabled=a[56];
			this.MnuFactDevVentas.Enabled=a[59];
			this.MnuFactNotaCV.Enabled=a[60];
			this.MnuFactDevCompras.Enabled=a[61];
			this.MnuFactNotaCC.Enabled=a[62];
			this.MnuFactProyectos.Enabled=a[63];
			this.MnuContMantTarj.Enabled=a[64];
			this.MnuContDepositoVouchers.Enabled=a[65];
			this.MnuContCuadreCaja.Enabled=a[66];
			this.MnuContFlujo.Enabled=a[67];
			this.MnuCont_Ind_Financieros.Enabled=a[68];
			this.MnuContPresupuesto.Enabled=a[69];
			this.mnuContLote.Enabled=a[70];
			this.mnuContCompara.Enabled=a[96];
			this.mnuContAsientoPeriodo.Enabled=a[97];
			this.mnuContLoteDepura.Enabled=a[98];
			this.mnuContLotAsiComprob.Enabled=a[99];
			this.mnuContLotCuentaObligatoria.Enabled=a[100];
			//			this.MnuContBloqAsientos.Enabled=a[71];
			this.MnuContCierPeriodo.Enabled=a[72];
			//this.mnuContRepAnt.Enabled = a[868];
			//			this.mnuContResCXP.Enabled = a[869];
			//			this.mnuContResCXC.Enabled = a[870];			
			this.mnuContabilidadAsDesc.Enabled = a[445];
			this.mnuContAntProv.Enabled = a[446];
			this.mnuContASF.Enabled = a[458];			
			this.mnuContMovBanc.Enabled = a[459];
			this.mnuContTEB.Enabled = a[467];
			this.mnuContPrestamos.Enabled = a[468];
			this.mnuContConsumosTC.Enabled = a[470];
			this.mnuContTCC.Enabled = a[486];
			this.mnuContRepGastos.Enabled = a[632];
			this.mnuContEstadoCheques.Enabled = a[687];
			this.mnuContRepFacElec.Enabled = a[688];
			this.mnuContBancosDPL.Enabled = a[705];
			this.mnuContBancosCDPL.Enabled = a[715];
			this.mnuContRB.Enabled = a[912];
			this.mnuContabilidadProvisiones.Enabled = a[913];
			this.mnuContRepEstadoCB.Enabled = a[921];
			this.mnuContCaja.Enabled = a[922];	
			this.mnuContEstadosCajas.Enabled = a[923];	
			this.mnuContIngresoDiarioCaja.Enabled = a[923];	
			this.mnuContIngresosDeCajas.Enabled = a[924];	
			this.mnuContEstadoGeneralDeCaja.Enabled = a[925];	
			this.mnuContBancosCBA.Enabled = a[919];	
			this.mnuContBancosCP.Enabled = a[920];	
			this.mnuContabilidadFPPP.Enabled = a[915];
			this.barButtonItem53.Enabled = a[916];
			this.mnuContRepPresupAlm.Enabled = a[927];

			

	

			#endregion Contabilidad
			#region Roles
			this.MnuRol.Enabled=a[5];
			this.MnuRolPersonal.Enabled=a[79];
			this.mnuRolPA.Enabled=a[79];
			this.MnuRolCargos.Enabled=a[80];
			this.MnuRolDpto.Enabled=a[81];
			this.MnuRolGenRol.Enabled=a[83];
			this.MnuRolRepRol.Enabled=a[85];
			this.MnuRolIncSuel.Enabled=a[87];
			this.MnuRolVerif.Enabled=a[82];
			this.mnuRolFiniquitos.Enabled=a[790];
			this.MnuRolBloqRol.Enabled=a[86];
			this.MnuRolAsigCtas.Enabled=a[90];
			this.barRolRubro.Enabled=a[84];
			this.mnuRolesLiquidacion.Enabled=a[88];
			this.mnuRolAnexo.Enabled=a[91];
			this.mnuRolUtilidad.Enabled=a[89];
			this.mnuTarea_Agr.Enabled = a[92];
			this.mnuRolesDistribucion.Enabled=a[93];
			this.mnuRolesAdicional.Enabled=a[94];
			this.mnuRolesAdicionalFormulas.Enabled=a[95];
			this.mnuRolPersonalPrueba.Enabled=a[732];
			this.mnuRolesPagoComisiones.Enabled=a[772];
			this.mnuRolesDecimoTercerSueldo.Enabled=a[972];
			this.mnuRolPorCajero.Enabled=a[1213];

			
			#endregion Roles
			#region Comisiones
			this.mnuRolesComisiones.Enabled=a[864];
			this.mnuRolComPIA.Enabled=a[865];
			this.mnuRolesCom_PA.Enabled=a[866];
			this.mnuRolesCom_PL.Enabled=a[867];
			this.mnuRolesCom_PML.Enabled=a[868];
			this.mnuRolesCom_AV.Enabled=a[869];
			this.mnuRolesComAJV.Enabled=a[875];
			this.mnuRolComisionesPC.Enabled=a[876];
			this.mnuRolCom_PorcCarpa.Enabled=a[877];
			this.mnuCom_PCGE.Enabled=a[878];
			this.mnuRolesCom_EPV.Enabled=a[879];
			this.mnuCom_Reportes.Enabled=a[880];
			this.mnuCom_ReporteLM.Enabled=a[1172];
			

			#endregion Comisiones
			#region Compras
			this.MnuCompras.Enabled=a[6];
			this.MnuFactProvee.Enabled=a[101];
			this.MnuFactOrdComp.Enabled=a[102];
			this.MnuFactCompras.Enabled=a[103];
			this.BarCompras.Enabled=a[103];
			this.MnuFactLotimp.Enabled=a[104];
			this.mnuFactVentasConsignacion.Enabled=a[105];
			this.MnuFactArtic.Enabled=a[106];
			this.BarArticulo.Enabled=a[106];
			this.MnuFactCombos.Enabled=a[107];
			this.MnnuFactLotesPromociones.Enabled=a[108];
			this.mnuFactActPrecio.Enabled=a[109];
			this.MnuFactLoteSpif.Enabled=a[110];
			this.mnuFactTransf.Enabled=a[111];
			this.MnuFactLoteTransMult.Enabled=a[112];
			this.MnuFactEstTransf.Enabled=a[113];			
			this.MnuFactLotesPromVen.Enabled=a[114];
			this.mnuComprasComboZOOM.Enabled=a[115];
			this.mnuComprasRepSpifVend.Enabled=a[461];
			this.mnuComprasReportes.Enabled = a[559];			
			this.mnuCompraReportesCPA.Enabled = a[560];
			this.mnuComprasGDA.Enabled = a[625];
			this.mnuComprasPromociones.Enabled = a[634];
			this.mnuComprasAPC.Enabled = a[663];
			this.mnuComprasProformas.Enabled = a[899];
			this.mnuCompraCSA.Enabled = a[986];
			this.mnuVentasReportesVA.Enabled = a[987];
			this.mnuComprasCatalogo.Enabled = a[988];
			this.mnuComprasAgenda.Enabled = a[989];
			this.mnuCompraGestion.Enabled = a[990];
			this.mnuComprasLiquidaGastos.Enabled = a[991];
			this.mnuMatrizDsctos.Enabled = a[992];
			this.mnuCompraCSA.Enabled = a[986];
			this.mnuVentasReportesVA.Enabled = a[987];
			this.mnuComprasCatalogo.Enabled = a[988];
			this.mnuComprasAgenda.Enabled = a[989];
			this.mnuCompraGestion.Enabled = a[990];
			this.mnuComprasLiquidaGastos.Enabled = a[991];
			this.mnuMatrizDsctos.Enabled = a[992];
			this.mnuCompraAODC.Enabled = a[1158];
			this.mnuComprasPromoFact.Enabled = a[1205];


			#endregion Compras
			#region Ventas
			this.mnuVentas.Enabled=a[7];
			this.MnuFactClientes.Enabled=a[116];
			this.BarCliente.Enabled=a[116];
			this.MnuFactVentas.Enabled=a[118];
			this.BarVentas.Enabled=a[118];
			this.MnuFactPreFactura.Enabled=a[119];
			this.MnuFactAnticipos.Enabled=a[120];
			this.MnuFactPedClientes.Enabled=a[121];
			this.MnuFactReservClientes.Enabled=a[122];
			this.MnuFactConsignacion.Enabled=a[123];
			this.mnuVentasComisiones.Enabled=a[128];
			this.MnuFacComConfiguracion.Enabled=a[129];
			this.MnuFactComCalificacion.Enabled=a[130];
			this.MnuFactDetalle.Enabled=a[131];
			this.MnuFactComConsulta.Enabled=a[132];
			this.mnuVentasResumenDiario.Enabled=a[127];
			this.mnuVentasCB.Enabled=a[125];
			this.mnuVentaAutDistribuidor.Enabled = a[126];
			this.mnuVentaLiqCons.Enabled = a[124];
			this.mnuVentasCotizacion.Enabled = a[117];
			this.mnuVentasImpresionDePrecios.Enabled = a[543];
			this.mnuVentasDA.Enabled = a[553];
//			this.mnuContEstadoCajas.Enabled = a[554];
			this.barSubItem3.Enabled = a[555];
			this.mnuVentasReportesPorMarca.Enabled = a[558];
			this.mnuVentasRVML.Enabled = a[570];
			this.mnuVentasRepMU.Enabled = a[589];
			this.mnuVentaRMU.Enabled = a[589];
			this.mnuVentasARP.Enabled = a[633];
			this.mnuVentasResumenComisiones.Enabled = a[658];
			this.mnuVentasIPP.Enabled = a[676];		
			this.mnuVentasMarcador.Enabled = a[678];	
			this.mnuVentasRepGE.Enabled = a[689];
			this.mnuVentasPlanifVentas.Enabled = a[719];
			this.mnuVentasPlanifJefeVentas.Enabled = a[724];
			this.barButtonItem66.Enabled = a[731];
			this.mnuVentasReportesDTC.Enabled = a[735];
			this.mnuVentasReportesEncuestas.Enabled = a[736];
			this.mnuVentasIPR.Enabled = a[763];
			this.mnuVentasPFR.Enabled = a[808];
			this.mnuVentasRP.Enabled = a[828];
			this.mnuVentasECW.Enabled = a[831];
			this.mnuVentasReportesVentasPorArticulos.Enabled = a[844];
			this.mnuVentasReportesVentasarticuloPorMes.Enabled = a[846];
			this.mnuComprasRepCodigoArticulos.Enabled = a[850];
			this.mnuVentasRPA.Enabled = a[886];
			this.mnuVentasCNOS.Enabled = a[897];
//			this.mnuVentasReportesIngresoDeCaja.Enabled = a[905];
//			this.mnuVentasReportesEstadosDeCajas.Enabled = a[910];
			this.mnuVentaRMUA.Enabled = a[911];
			
			
			#endregion Ventas
			#region CrediPoint
			this.MnuCredito.Enabled = a[8];
			this.MnuCreditoSolRapida.Enabled = a[135];
			this.mnuCrediPointSolicitud.Enabled = a[136];
			this.mnuCrediPointMantenedor.Enabled = a[137];
			this.mnuCrediPointVerTerreno.Enabled = a[138];
			this.mnuCrediPointGestiones.Enabled = a[139];
			this.mnuCrediPointGestVerCed.Enabled = a[422];
			this.mnuCrediPointGestRevCed.Enabled = a[423];
			this.mnuContRepAnt.Enabled = a[427];
			this.mnuCrediPointCatalogos.Enabled = a[140];
			this.MnuCreCobranza.Enabled = a[133];
			this.mnuCreLotesUnibanco.Enabled = a[141];
			this.MnuCreditoCotizador.Enabled = a[134];
			this.MnuCreditoCuota.Enabled = a[142];
			this.mnuContCCxCCP.Enabled = a[436];
			this.mnuCPReportes.Enabled = a[562];
			this.mnuCPReportesCliPoint.Enabled = a[563];
			this.mnuCPInformeCobrador.Enabled = a[564];
			this.mnuCrediPointDFG.Enabled = a[574];
			this.mnuCrediPointMEF.Enabled = a[575];
			this.mnuCPCP.Enabled = a[577];
			this.mnuCPIO.Enabled = a[578];
			this.mnuCPLCS.Enabled = a[585];
			this.mnuCPRepEquifax.Enabled = a[586];
			this.mnuCPMFCP.Enabled = a[587];
			this.mnuCPRRCS.Enabled = a[624];
			this.mnuCPRD.Enabled = a[629];
			this.mnuCPVEP.Enabled = a[630];
			this.mnuCPRepECP.Enabled = a[631];
			this.mnuCPEDC.Enabled = a[635];
			this.mnuCPHRC.Enabled = a[637];
			this.mnuCPRRC.Enabled = a[639];
			this.mnuCPReportesRCV.Enabled = a[657];
			this.mnuCPDFGP.Enabled = a[662];
			this.mnuCPCobranzas.Enabled = a[679];
			this.mnuVentasGCCC.Enabled = a[753];
			this.mnuCPCobranzasLPCA.Enabled = a[766];
			this.mnuCPReportesEC.Enabled = a[809];
			this.mnuCPCobranzasEDC.Enabled = a[861];
			this.mnuCPCartera.Enabled = a[863];
			this.mnuCobranzasRGC.Enabled = a[894];		

			this.mnuCPRepCP.Enabled = a[932];
			this.mnuCPRepRDD.Enabled = a[933];	
			this.mnuCPRepECP.Enabled = a[934];	
			this.barButtonItem37.Enabled = a[935];
			this.mnuCPReportesGDCM.Enabled = a[936];
			this.mnuCPReportesAbonos.Enabled = a[937];	
			this.mnuCPReportesNDC.Enabled = a[938];
			this.mnuCPReportesBR.Enabled = a[939];
			this.barButtonItem76.Enabled = a[940];
			this.mnuCPReportesREC.Enabled = a[941];
			this.mnuCrediPOINTReportesRMR.Enabled = a[942];
			this.mnuCPReportesRRM.Enabled = a[943];
			this.mnuCPReportesCVPM.Enabled = a[944];
			this.mnuCPCobranzasGDC.Enabled = a[945];
			this.mnuCPInformeCobrador.Enabled = a[946];
			this.mnuCPIO.Enabled = a[947];
			this.mnuCPCP.Enabled = a[948];
			this.mnuCPCobranzasRCP.Enabled = a[949];
			this.barButtonItem50.Enabled = a[950];
			this.mnuCPCobranzasCCPPC.Enabled = a[951];
			this.barButtonItem72.Enabled = a[952];
			this.mnuCPCobranzasMRPR.Enabled = a[953];
			this.barSubItem10.Enabled = a[954];
			this.mnuCrediPointDFG.Enabled = a[955];		
			this.mnuCreLotesUnibanco.Enabled = a[956];
			this.mnuContRepAnt.Enabled = a[957];	
			this.mnuCPRRCS.Enabled = a[958];
			this.mnuContCCxCCP.Enabled = a[959];	
			this.barSubItem11.Enabled = a[960];	
			this.mnuCpLbp.Enabled = a[961];	
			this.mnuCPLotesPPagos.Enabled = a[962];	
			this.mnuCPDFGP.Enabled = a[963];	
			this.mnuCrediPOINTRMR.Enabled = a[964];	
			this.mnuCPCobranzasCarteraAIP.Enabled = a[969];	
			this.mnuCPCobranzasRPE.Enabled = a[970];
			this.mnuCPCRD.Enabled = a[980];	
			this.mnuCPCDRDC.Enabled = a[981];	
			this.mnuCPCDRVC.Enabled = a[982];	
			this.mnuCPCDSR.Enabled = a[983];	
			this.mnuCPCDRSLT.Enabled = a[984];	
			this.mnuCPCDLCPG.Enabled = a[985];	
			this.mnuCRResumenCredito.Enabled = a[1182];
			this.barButtonItem82.Enabled = a[1184];
      this.mnuCPComisiones.Enabled = a[1188];
			 this.mnuCPCobranzasCboOpe.Enabled = a[1194];
			this.mnuCPCobranzasCERTIFICADO.Enabled = a[1197];
			this.mnuRPtComprasPuC.Enabled = a[1198];
			this.mnuCreditoProReal.Enabled = a[1193];


			this.mnuCPComisionesAsigCobr.Enabled = a[1189];
			this.mnuCPComisionesTrmOpe.Enabled = a[1190];
			this.mnuCPComisionesCom.Enabled = a[1191];
			this.mnuCreditoGesCall.Enabled = a[1192];
			this.barButtonItem84.Enabled = a[1212];

			
	
			#endregion CrediPoint
			#region Produccion
			this.MnuProduc.Enabled=a[9];
			this.MnuProducRecetas.Enabled=a[144];
			this.MnuProducOrdProd.Enabled=a[145];
			this.MnuProdEstado.Enabled=a[146];
			this.mnuProduccionPendientes.Enabled=a[147];
			#endregion Produccion
			#region Departamento Técnico
			this.mnuDT.Enabled=a[10];
			this.MnuDTSTA.Enabled=a[148];
			this.MnuDTSolicitud.Enabled=a[149];
			this.mnuDTIngresoMultiple.Enabled=a[525];
			this.mnuDTRetiro.Enabled=a[533];
			this.mnuDTIRM.Enabled=a[640];
			this.mnuDTIRMP.Enabled=a[641];
			this.mnuDTProductosDeFeria.Enabled = a[722];
			this.mnuDTIRAT.Enabled=a[737];
			this.mnuDTIREP.Enabled=a[738];
			this.mnuDTIRRPT.Enabled=a[739];
			this.mnuDTIREPC.Enabled=a[740];
			this.mnuDTIRPPL.Enabled=a[741];
			this.mnuDTIRAPD.Enabled=a[742];
			this.mnuDTIRSNC.Enabled=a[743];
			this.mnuDTIRREP.Enabled=a[744];
			this.mnuDTIRRMA.Enabled=a[747];
			this.mnuDTIRFDT.Enabled=a[750];
			this.mnuDTIRSIN.Enabled=a[751];
			this.mnuDTIRCDE.Enabled=a[752];
			this.mnuDTIRRRP.Enabled=a[745];
			this.mnuDTIRRTE.Enabled=a[746];
			this.mnuDTIRRPG.Enabled=a[748];
			this.mnuDTIRRSR.Enabled=a[749];
			this.mnuComprasRepCodigoArticulos.Enabled=a[848];
			#endregion Departamento Técnico			
			#region Lote
			this.mnuLotte.Enabled = a[11];
			this.mnuHabitaciones.Enabled = a[593];
			this.mnuTmaHabitacion.Enabled = a[594];
			this.mnuTmpEstadia.Enabled = a[595];
			this.mnuTmpHabitaciones.Enabled = a[596];
			this.mnuCuadroHabitaciones.Enabled = a[597];
			this.mnuCobrPrenda.Enabled = a[598];
			this.mnuKardexLot.Enabled=a[599];
			this.mnuPromociones.Enabled = a[600];
			this.mnuReservas.Enabled=a[601];
			this.mnuPSRDC.Enabled=a[833];
			this.mnuTarjetasDeIngreso.Enabled=a[895];

			this.barButtonItem27.Enabled=a[997];
			this.barButtonItem28.Enabled=a[998];
			this.barButtonItem29.Enabled=a[999];
			this.barButtonItem36.Enabled=a[1000];
			this.barSubItem9.Enabled=a[1001];
			this.barButtonItem30.Enabled=a[1002];
			this.mnuLottePlanificacion.Enabled=a[1003];
			this.barButtonItem73.Enabled=a[1004];
			this.mnuLotteCompetencia.Enabled=a[1005];

			this.barButtonItem24.Enabled=a[1006];
			this.barButtonItem25.Enabled=a[1007];
			this.barButtonItem31.Enabled=a[1008];
			this.barButtonItem32.Enabled=a[1009];
			this.barButtonItem35.Enabled=a[1011];
			this.barButtonItem40.Enabled=a[1012];
			this.barButtonItem42.Enabled=a[1013];
			this.barButtonItem67.Enabled=a[1014];
			this.barSubItem14.Enabled=a[1015];
			this.barButtonItem75.Enabled=a[1016];
			this.mnuLotteReportesIT.Enabled=a[1017];
			this.mnuLotteReporteIPH.Enabled=a[1018];
			this.mnuLoteRepProcSIM.Enabled=a[1019];
			this.barButtonItem69.Enabled=a[1020];
      





			#endregion Lote
      #region Reporteria

			this.barSubItem5.Enabled=a[1021];
      this.mnuRPtBodega.Enabled=a[1022];
			               this.mnuRPtBodegaInvFEcha.Enabled=a[569];
			               this.mnuRPtBodegaRptTrans.Enabled=a[684];
			               this.mnuRPtBodegaInvVal.Enabled=a[685];
			               this.mnuRPtBodegaKXSer.Enabled=a[1030];
			               this.mnuRPtBodegaEsConInv.Enabled=a[1031];
			               this.mnuRPtBodegaMovDesc.Enabled=a[1032];
			                 this.mnuRPtBodegaKCC.Enabled=a[1146];
			

	
			this.mnuRPtContabilidad.Enabled=a[1023];
							
						this.mnuRPtContabilidadCont.Enabled=a[45];
						this.mnuRPtContabilidadRet.Enabled=a[75];
						this.mnuRPtContabilidadAsDEsc.Enabled=a[445];
						this.mnuRPtContabilidadGast.Enabled=a[632];
						this.mnuRPtContabilidadFactEl.Enabled=a[688];
						this.mnuRPtContabilidadCaCh.Enabled=a[917];
						this.mnuRPtContabilidadGaAn.Enabled=a[918];
						this.mnuRPtContabilidadEstaConB.Enabled=a[921];
						this.mnuRPtContabilidadPrXAlm.Enabled=a[927];
            this.mnuRPPagoAgil.Enabled=a[1195];
			      this.mnuRPtContabilidadSCosecha.Enabled=a[1233];

			
			      
			     
			this.mnuRPtRoles.Enabled=a[1024];
									this.mnuRPtRolesPFA.Enabled=a[1033];
									this.mnuRPtRolesVeXM.Enabled=a[1034];
									this.mnuRPtRolesUni.Enabled=a[1035];
									this.mnuRPtRolesNG.Enabled=a[1036];
									this.mnuRPtRolesAEM.Enabled=a[1037];
									this.mnuRPtRolesHHT.Enabled=a[1038];
									this.mnuRPtRolesIXD.Enabled=a[1039];
   
			this.mnuRPtCompras.Enabled=a[1025];
									 this.mnuRPtComprasCPA.Enabled=a[560];
									 this.mnuRPtComprasT.Enabled=a[561];
									 this.mnuRPtComprasRAC.Enabled=a[663];
									 this.mnuRPtComprasCA.Enabled=a[850];

			this.mnuRPtVentas.Enabled=a[1026];

			     
//						this.mnuRPTVentasPorcMarca.Enabled=a[1041];
//						this.mnuRPTVentasPorcMarcaUtilid.Enabled=a[1042];
//						this.mnuRPTVentasMargUti.Enabled= a[1043];
//						this.mnuRPTVentasMaUXArt.Enabled= a[1044];
//						this.mnuRPTVentasAnInMEn.Enabled= a[1045];
//						this.mnuRPTVentasRevFact.Enabled= a[1046];
//						this.mnuRPTVentasXDia.Enabled= a[1047];
//						this.mnuRPTVentasNFactPForPago.Enabled= a[1048];
//						this.mnuRPTVentasGaraExt.Enabled=a[1049];
//						this.mnuRPTVentasXFecha.Enabled=a[1050];
//						this.mnuRPTVentasComDeVentas.Enabled=a[1051];
//						this.mnuRPTVentasANuXLo.Enabled=a[1052];
//						this.mnuRPTVentasMenXLo.Enabled=a[1053];
//						this.mnuRPTVentasTarCredi.Enabled=a[1054];
//						this.mnuRPTVentasDetTc.Enabled=a[1055];
//						this.mnuRPTVentasEncuestas.Enabled=a[1056];
//						this.mnuRPTVentasReCXCo.Enabled=a[1057];
//						this.mnuRPTVentasXano.Enabled=a[1058];
//						this.mnuRPTVentasXArtic.Enabled=a[1059];
//						this.mnuRPTVentasArtXMes.Enabled=a[1060];
//						this.barSubItem12.Enabled=a[1151];
//						//			if (!a[1026]) this.mnuRPTVentasComisionesConf.Enabled=a[1041];
//						//			if (!a[1026]) this.mnuRPTVentasComisionesCali.Enabled=a[1041];
//						//			if (!a[1026]) this.mnuRPTVentasComisionesDetGe.Enabled=a[1041];
//						//			if (!a[1026]) this.mnuRPTVentasComisionesConsulta.Enabled=a[1041];


		       
		       	this.mnuReporVentasDFI.Enabled=a[1231];

			     this.mnuRPTVentasPorcMargUtilidad.Enabled=a[1144];
			     this.mnuReporVentasPDGA.Enabled=a[1174];
			this.mnuRPtCREPOINT.Enabled=a[1027];
					#region credipoint
						this.mnuRPtCREPOINTFDG.Enabled=a[1226];
			      this.mnuRPtCREPOINRMD.Enabled=a[1228];

					#endregion credipoint
			   
			this.mnuRPtDT.Enabled=a[1028];
			this.mnuRPtPeonyServicios.Enabled=a[1029];


		


			#endregion Reporteria

			this.mnuMetricas.Enabled =a[821];
			this.mnuPCGEx.Enabled =a[822];
			this.mnuSTA.Enabled =a[823];
			this.mnuIngreso.Enabled =a[1143];
			this.mnuCPCobCM.Enabled = a[1176];
			
		}

		/// <summary>
		/// Visibilidad de Menus
		/// </summary>
		/// <param name="a"></param>
		private void AsignaVectorV(bool [] a)
		{
			#region Archivo
			if (!a[1]) this.MnuArch.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[11])
			{
				//				this.MnuArchNuevo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				//				this.barNuevo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}
			if (!a[12])
			{
				this.MnuArchAbrir.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.barAbrir.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}
			if (!a[13])
			{
				this.MnuArchGuardar.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.barGrabar.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}
			if (!a[14]) this.MnuArchImpoExpo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[15]) this.MnuArchProp.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[16]) this.MnuArchConfImp.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[17]) this.MnuArchImpoLat.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[18]) this.MnuArchImpoOtros.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[19]) this.mnuArchPropGen.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[20]) this.MnuArchPropCont.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[21]) this.MnuArchPropFact.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[22]) this.MnuArchPropRol.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[23]) this.mnuArchPropEmpresa.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[24]) this.mnuArchPropConfig.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			#endregion Archivo
			#region Herramientas
			if (!a[2]) this.MnuHerr_.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[25]) this.MnuHerr_Camb_Contr.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[26]) this.MnuHerr_Actu_Vers.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[27]) this.MnuHerr_Sal_Ini.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[28]) this.MnuHerr_Camb_Ctes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[29]) this.MnuHerr_Copi_Arch.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[30]) this.MnuHerr_Auditoría.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[31]) this.MnuHerr_Act_Saldo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[32]) this.MnuEdicSeguridad.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[33])
			{
				this.MnuEdicSegMenu.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.barSeguridad.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}
			if (!a[34]) this.MnuEdicSegUsuario.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[433]) this.mnuHerramientasNumeracion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[544]) this.mnuHerrSegCD.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[568]) this.mnuHerramientasCRIVA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[602]) this.mnuHerrALF.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[733]) this.mnuHerramientasPTP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[734]) this.mnuHerramientasAP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
      if (!a[908]) this.mnuHerramientasPTA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
      if (!a[928]) this.mnuHerramientasAPL.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[929]) this.mnuHerramientasASDP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			

			#endregion Herramientas
			#region Bodega
			if (!a[3]) this.MnuBodega.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[881]) this.mnuBodegaTRM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[35]) this.MnuFactSerie.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[36]) this.MnuFacSeriales.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[37]) this.MnuFacInv.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[38]) this.mnuBodegaKardex.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
			if (!a[39]) this.mnuBodegaGRC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[40]) this.MnuFactEnvTransferencia.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[41]) this.MnuFactRecTransferencia.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[42]) this.mnuFactBodegas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[541]) this.mnuBodegaCI.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[542]) this.mnuBodegaDCI.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[551]) this.mnuBodegaCierreInventario.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[569]) this.mnuBodegaIF.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[685]) this.mnuBodegaInvVal.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[579]) this.mnuBodegaSectorizacion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[603]) this.mnuBodegaIngSeriales.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[682]) this.mnuComprasTEL.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[683]) this.mnuReportesBodega.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[684]) this.mnuRepCmpTransf.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[696]) this.mnuBodegaIPR.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			#endregion Bodega
			#region Contabilidad
			if (!a[4]) this.MnuCont.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[43])
			{
				this.MnuContPlanCtas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.BarPlanCuentas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}			
			if (!a[44])
			{
				this.MnuContAsientos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.barAsiento.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}
			if (!a[45])
			{
				this.MnuContRepCont.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.BarRepContabilidad.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}
			if (!a[46]) this.MnuContAsigCtas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[47]) this.mnuContReportes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[48]) this.MnuContTransacciones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[73]) this.MnuContTransE.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[74]) this.MnuContTransI.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[75]) this.MnuContTransR.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[49]) this.MnuContSRI.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[76]) this.mnuContTransAnexoTr.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[77]) this.mnuContTransFormulario.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[78]) this.mnuContTransConciliacion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[50])
			{
				this.MnuContCXP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.barPagar.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}
			if (!a[51]) this.MnuFactAutPagos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[52]) this.MnuFacPagosAut.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[53])
			{
				this.MnuContCXC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.barCobro.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}
			if (!a[54]) this.MnuFacArtServicios.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[55]) this.MnuFactActFijos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[56]) this.MnuContBcos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[57]) this.MnuContCtas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[58]) this.MnuContEstadoCta.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
			if (!a[59]) this.MnuFactDevVentas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[60]) this.MnuFactNotaCV.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[61]) this.MnuFactDevCompras.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[62]) this.MnuFactNotaCC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[63]) this.MnuFactProyectos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[64]) this.MnuContMantTarj.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[65]) this.MnuContDepositoVouchers.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[66]) this.MnuContCuadreCaja.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[67]) this.MnuContFlujo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[68]) this.MnuCont_Ind_Financieros.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[69]) this.MnuContPresupuesto.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[70]) this.mnuContLote.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			//			if (!a[71]) this.MnuContBloqAsientos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[72]) this.MnuContCierPeriodo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
			if (!a[96]) this.mnuContCompara.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[97]) this.mnuContAsientoPeriodo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[98]) this.mnuContLoteDepura.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[99]) this.mnuContLotAsiComprob.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[100]) this.mnuContLotCuentaObligatoria.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			//if (!a[868]) this.mnuContRepAnt.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			//			if (!a[869]) this.mnuContResCXP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			//			if (!a[870]) this.mnuContResCXC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
			if (!a[445]) this.mnuContabilidadAsDesc.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[446]) this.mnuContAntProv.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[458]) this.mnuContASF.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[459]) this.mnuContMovBanc.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[467]) this.mnuContTEB.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;	
			if (!a[468]) this.mnuContPrestamos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;				
			if (!a[470]) this.mnuContConsumosTC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;				
			if (!a[486]) this.mnuContTCC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;				
			if (!a[632]) this.mnuContRepGastos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;							
			if (!a[687]) this.mnuContEstadoCheques.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[688]) this.mnuContRepFacElec.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;				
			if (!a[705]) this.mnuContBancosDPL.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[715]) this.mnuContBancosCDPL.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				if (!a[912]) this.mnuContRB.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;	
			if (!a[913]) this.mnuContabilidadProvisiones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;	
      if (!a[921]) this.mnuContRepEstadoCB.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;	
			 if (!a[922]) this.mnuContCaja.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[923]) this.mnuContEstadosCajas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[924]) this.mnuContIngresoDiarioCaja.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[925]) this.mnuContIngresosDeCajas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[926]) this.mnuContEstadoGeneralDeCaja.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[919]) this.mnuContBancosCBA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[920]) this.mnuContBancosCP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[915]) this.mnuContabilidadFPPP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[916]) this.barButtonItem53.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[927]) this.mnuContRepPresupAlm.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			
			#endregion Contabilidad			
			#region Roles
			if (!a[5]) this.MnuRol.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[79]) 
			{
				this.MnuRolPersonal.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.mnuRolPA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}

			if (!a[80]) this.MnuRolCargos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[81]) this.MnuRolDpto.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[83]) this.MnuRolGenRol.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[85]) this.MnuRolRepRol.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[87]) this.MnuRolIncSuel.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[82]) this.MnuRolVerif.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[790]) this.mnuRolFiniquitos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never; //validar 
			if (!a[86]) this.MnuRolBloqRol.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[90])	this.MnuRolAsigCtas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;				
			if (!a[84]) this.barRolRubro.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[88]) this.mnuRolesLiquidacion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[91]) this.mnuRolAnexo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[89]) this.mnuRolUtilidad.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[92]) this.mnuTarea_Agr.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[93]) this.mnuRolesDistribucion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[94]) this.mnuRolesAdicional.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[95]) this.mnuRolesAdicionalFormulas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[732]) this.mnuRolPersonalPrueba.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[772]) this.mnuRolesPagoComisiones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[972]) this.mnuRolesDecimoTercerSueldo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1213]) this.mnuRolPorCajero.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

		

		
			#endregion Roles
			#region Comisiones
			if (!a[864]) this.mnuRolesComisiones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[865]) this.mnuRolComPIA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[866]) this.mnuRolesCom_PA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[867]) this.mnuRolesCom_PL.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[868]) this.mnuRolesCom_PML.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[869]) this.mnuRolesCom_AV.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[875]) this.mnuRolesComAJV.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			if (!a[876]) this.mnuRolComisionesPC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[877]) this.mnuRolCom_PorcCarpa.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[878]) this.mnuCom_PCGE.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			if (!a[879]) this.mnuRolesCom_EPV.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[880]) this.mnuCom_Reportes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;	
			if (!a[1172]) this.mnuCom_ReporteLM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;	
			
			#endregion Comisiones
			#region Compras
			if (!a[6]) this.MnuCompras.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[101]) this.MnuFactProvee.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[102]) this.MnuFactOrdComp.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[103])
			{
				this.MnuFactCompras.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.BarCompras.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}		
			if (!a[104]) this.MnuFactLotimp.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[105]) this.mnuFactVentasConsignacion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[106])
			{
				this.MnuFactArtic.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.BarArticulo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}			
			if (!a[107]) this.MnuFactCombos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[108]) this.MnnuFactLotesPromociones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[109]) this.mnuFactActPrecio.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[111]) this.mnuFactTransf.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[112]) this.MnuFactLoteTransMult.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[113]) this.MnuFactEstTransf.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
			if (!a[110]) this.MnuFactLoteSpif.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[114]) this.MnuFactLotesPromVen.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
			if (!a[115]) this.mnuComprasComboZOOM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
			if (!a[461]) this.mnuComprasRepSpifVend.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
			if (!a[559]) this.mnuComprasReportes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;						
			if (!a[560]) this.mnuCompraReportesCPA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[625]) this.mnuComprasGDA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[634]) this.mnuComprasPromociones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[663]) this.mnuComprasAPC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[850]) this.mnuComprasRepCodigoArticulos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[899]) this.mnuComprasProformas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			
			if (!a[986]) this.mnuCompraCSA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[987]) this.mnuVentasReportesVA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[988]) this.mnuComprasCatalogo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[989]) this.mnuComprasAgenda.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[990]) this.mnuCompraGestion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[991]) this.mnuComprasLiquidaGastos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[992]) this.mnuMatrizDsctos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			
			if (!a[986]) this.mnuCompraCSA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[987]) this.mnuVentasReportesVA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[988]) this.mnuComprasCatalogo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[989]) this.mnuComprasAgenda.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[990]) this.mnuCompraGestion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[991]) this.mnuComprasLiquidaGastos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[992]) this.mnuMatrizDsctos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1158]) this.mnuCompraAODC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1205]) this.mnuComprasPromoFact.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			#endregion Compras
			#region Ventas
			if (!a[7]) this.mnuVentas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[116])
			{
				this.MnuFactClientes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.BarCliente.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			} 
			if (!a[117]) this.mnuVentasCotizacion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[118])
			{
				this.MnuFactVentas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				this.BarVentas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			}
			if (!a[119]) this.MnuFactPreFactura.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[120]) this.MnuFactAnticipos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[121]) this.MnuFactPedClientes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[122]) this.MnuFactReservClientes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[123]) this.MnuFactConsignacion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[124]) this.mnuVentaLiqCons.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[128]) this.mnuVentasComisiones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[129]) this.MnuFacComConfiguracion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[130]) this.MnuFactComCalificacion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[131]) this.MnuFactDetalle.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[132]) this.MnuFactComConsulta.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[127]) this.mnuVentasResumenDiario.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[125]) this.mnuVentasCB.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[126]) this.mnuVentaAutDistribuidor.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[543]) this.mnuVentasImpresionDePrecios.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[553]) this.mnuVentasDA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//			if (!a[554]) this.mnuContEstadoCajas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[555]) this.barSubItem3.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[558]) this.mnuVentasReportesPorMarca.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[570]) this.mnuVentasRVML.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[589]) this.mnuVentasRepMU.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[589]) this.mnuVentaRMU.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[633]) this.mnuVentasARP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;	
			if (!a[658]) this.mnuVentasResumenComisiones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[676]) this.mnuVentasIPP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
			if (!a[678]) this.mnuVentasMarcador.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[689]) this.mnuVentasRepGE.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[719]) this.mnuVentasPlanifVentas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[724]) this.mnuVentasPlanifJefeVentas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[731]) this.barButtonItem66.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[735]) this.mnuVentasReportesDTC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[736]) this.mnuVentasReportesEncuestas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[763]) this.mnuVentasIPR.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[808]) this.mnuVentasPFR.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[828]) this.mnuVentasRP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[831]) this.mnuVentasECW.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[844]) this.mnuVentasReportesVentasPorArticulos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			//if (!a[856]) this.mnuVentasReportesVentasPorArticulos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[886]) this.mnuVentasRPA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[897]) this.mnuVentasCNOS.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//			if (!a[905]) this.mnuVentasReportesIngresoDeCaja.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//      if (!a[910]) this.mnuVentasReportesEstadosDeCajas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[911]) this.mnuVentaRMUA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			
			
			#endregion Ventas
			#region CrediPoint
			if (!a[8]) this.MnuCredito.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[133]) this.MnuCreCobranza.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
			if (!a[134]) this.MnuCreditoCotizador.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[135]) this.MnuCreditoSolRapida.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[136]) this.mnuCrediPointSolicitud.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[137]) this.mnuCrediPointMantenedor.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[138]) this.mnuCrediPointVerTerreno.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[139]) this.mnuCrediPointGestiones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[140]) this.mnuCrediPointCatalogos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[141]) this.mnuCreLotesUnibanco.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[142]) this.MnuCreditoCuota.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[422]) this.mnuCrediPointGestVerCed.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[423]) this.mnuCrediPointGestRevCed.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[427]) this.mnuContRepAnt.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[436]) this.mnuContCCxCCP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[562]) this.mnuCPReportes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[563]) this.mnuCPReportesCliPoint.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[564]) this.mnuCPInformeCobrador.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[574]) this.mnuCrediPointDFG.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[575]) this.mnuCrediPointMEF.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[577]) this.mnuCPCP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[578]) this.mnuCPIO.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[585]) this.mnuCPLCS.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[586]) this.mnuCPRepEquifax.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[587]) this.mnuCPMFCP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[624]) this.mnuCPRRCS.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[629]) this.mnuCPRD.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[630]) this.mnuCPVEP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[631]) this.mnuCPRepECP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[635]) this.mnuCPEDC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[637]) this.mnuCPHRC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[639]) this.mnuCPRRC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[657]) this.mnuCPReportesRCV.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[662]) this.mnuCPDFGP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[679]) this.mnuCPCobranzas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[753]) this.mnuVentasGCCC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[766]) this.mnuCPCobranzasLPCA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[809]) this.mnuCPReportesEC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;	
			if (!a[861]) this.mnuCPCobranzasEDC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;	
			if (!a[863]) this.mnuCPCartera.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;	
			if (!a[894]) this.mnuCobranzasRGC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;		
		
			
			if (!a[932]) this.mnuCPRepCP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[933]) this.mnuCPRepRDD.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[934]) this.mnuCPRepECP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[935]) this.barButtonItem37.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[936]) this.mnuCPReportesGDCM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[937]) this.mnuCPReportesAbonos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[938]) this.mnuCPReportesNDC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[939]) this.mnuCPReportesBR.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[940]) this.barButtonItem76.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[941]) this.mnuCPReportesREC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[942]) this.mnuCrediPOINTReportesRMR.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[943]) this.mnuCPReportesRRM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[944]) this.mnuCPReportesCVPM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[945]) this.mnuCPCobranzasGDC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[946]) this.mnuCPInformeCobrador.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[947]) this.mnuCPIO.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[948]) this.mnuCPCP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[949]) this.mnuCPCobranzasRCP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[950]) this.barButtonItem50.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[951]) this.mnuCPCobranzasCCPPC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[952]) this.barButtonItem72.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[953]) this.mnuCPCobranzasMRPR.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[954]) this.barSubItem10.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[955]) this.mnuCrediPointDFG.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[956]) this.mnuCreLotesUnibanco.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[957]) this.mnuContRepAnt.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[958]) this.mnuCPRRCS.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[959]) this.mnuContCCxCCP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[960]) this.barSubItem11.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[961]) this.mnuCpLbp.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[962]) this.mnuCPLotesPPagos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[963]) this.mnuCPDFGP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[964]) this.mnuCrediPOINTRMR.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[969]) this.mnuCPCobranzasCarteraAIP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[970]) this.mnuCPCobranzasRPE.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			if (!a[980]) this.mnuCPCRD.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[981]) this.mnuCPCDRDC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[982]) this.mnuCPCDRVC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[983]) this.mnuCPCDSR.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[984]) this.mnuCPCDRSLT.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[985]) this.mnuCPCDLCPG.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1182])this.mnuCRResumenCredito.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1184])this.barButtonItem82.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1188])this.mnuCPComisiones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1194])this.mnuCPCobranzasCboOpe.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1197]) this.mnuCPCobranzasCERTIFICADO.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1198]) this.mnuRPtComprasPuC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1193]) this.mnuCreditoProReal.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			if (!a[1189]) this.mnuCPComisionesAsigCobr.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1190]) this.mnuCPComisionesTrmOpe.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1191]) this.mnuCPComisionesCom.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1192]) this.mnuCreditoGesCall.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1212]) this.barButtonItem84.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
     


		

			

			#endregion CrediPoint
			#region Produccion
			if (!a[9]) this.MnuProduc.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[144]) this.MnuProducRecetas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[145]) this.MnuProducOrdProd.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[146]) this.MnuProdEstado.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[147]) this.mnuProduccionPendientes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			#endregion Produccion
			#region Departamento Técnico
			if (!a[10]) this.mnuDT.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[148]) this.MnuDTSTA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[149]) this.MnuDTSolicitud.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[525]) this.mnuDTIngresoMultiple.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[533]) this.mnuDTRetiro.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[640]) this.mnuDTIRM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
			if (!a[641]) this.mnuDTIRMP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[722]) this.mnuDTProductosDeFeria.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[737]) this.mnuDTIRAT.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[738]) this.mnuDTIREP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[739]) this.mnuDTIRRPT.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[740]) this.mnuDTIREPC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[741]) this.mnuDTIRPPL.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[742]) this.mnuDTIRAPD.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[743]) this.mnuDTIRSNC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[744]) this.mnuDTIRREP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[747]) this.mnuDTIRRMA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[750]) this.mnuDTIRFDT.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[751]) this.mnuDTIRSIN.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[752]) this.mnuDTIRCDE.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[745]) this.mnuDTIRRRP.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[746]) this.mnuDTIRRTE.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[748]) this.mnuDTIRRPG.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[749]) this.mnuDTIRRSR.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[848]) this.mnuComprasRepCodigoArticulos.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

	
			#endregion Departamento Técnico				
			#region Lotte
			if (!a[11]) this.mnuLotte.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[593]) this.mnuHabitaciones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[594]) this.mnuTmaHabitacion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[595]) this.mnuTmpEstadia.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[596]) this.mnuTmpHabitaciones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[597]) this.mnuCuadroHabitaciones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[598]) this.mnuCobrPrenda.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[599]) this.mnuKardexLot.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[600]) this.mnuPromociones.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[601]) this.mnuReservas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[833]) this.mnuPSRDC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[895]) this.mnuTarjetasDeIngreso.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			if (!a[997]) this.barButtonItem27.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[998]) this.barButtonItem28.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[999]) this.barButtonItem29.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1000]) this.barButtonItem36.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1001]) this.barSubItem9.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1002]) this.barButtonItem30.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1003]) this.mnuLottePlanificacion.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1004]) this.barButtonItem73.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1005]) this.mnuLotteCompetencia.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			if (!a[1006]) this.barButtonItem24.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1007]) this.barButtonItem25.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1008]) this.barButtonItem31.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1009]) this.barButtonItem32.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1011]) this.barButtonItem35.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1012]) this.barButtonItem40.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1013]) this.barButtonItem42.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1014]) this.barButtonItem67.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1015]) this.barSubItem14.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1016]) this.barButtonItem75.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1017]) this.mnuLotteReportesIT.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1018]) this.mnuLotteReporteIPH.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1019]) this.mnuLoteRepProcSIM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1020]) this.barButtonItem69.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			#endregion Lotte
			#region Directorio 
			if (!a[898]) this.mnuDirectorio.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			#endregion Directorio 

			#region Reporteria
			if (!a[1021]) this.barSubItem5.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
					if (!a[1022]) this.mnuRPtBodega.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[569]) this.mnuRPtBodegaInvFEcha.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[684]) this.mnuRPtBodegaRptTrans.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[685]) this.mnuRPtBodegaInvVal.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[1030]) this.mnuRPtBodegaKXSer.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[1031]) this.mnuRPtBodegaEsConInv.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[1032]) this.mnuRPtBodegaMovDesc.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
		            	if (!a[1146]) this.mnuRPtBodegaKCC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			
			    
					if (!a[1023]) this.mnuRPtContabilidad.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			               
								  if (!a[45]) this.mnuRPtContabilidadCont.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[75]) this.mnuRPtContabilidadRet.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[445]) this.mnuRPtContabilidadAsDEsc.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[632]) this.mnuRPtContabilidadGast.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[688]) this.mnuRPtContabilidadFactEl.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[917]) this.mnuRPtContabilidadCaCh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[918]) this.mnuRPtContabilidadGaAn.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
							  	if (!a[921]) this.mnuRPtContabilidadEstaConB.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
									if (!a[927]) this.mnuRPtContabilidadPrXAlm.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			            if (!a[1195]) this.mnuRPPagoAgil.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			            if (!a[1233]) this.mnuRPtContabilidadSCosecha.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

 
					if (!a[1024]) this.mnuRPtRoles.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

												if (!a[1033]) this.mnuRPtRolesPFA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
												if (!a[1034]) this.mnuRPtRolesVeXM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
												if (!a[1035]) this.mnuRPtRolesUni.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
												if (!a[1036]) this.mnuRPtRolesNG.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
												if (!a[1037]) this.mnuRPtRolesAEM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
										  	if (!a[1038]) this.mnuRPtRolesHHT.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
												if (!a[1039]) this.mnuRPtRolesIXD.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

   
					if (!a[1025]) this.mnuRPtCompras.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
												if (!a[560]) this.mnuRPtComprasCPA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
												if (!a[561]) this.mnuRPtComprasT.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
										  	if (!a[663]) this.mnuRPtComprasRAC.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
												if (!a[850]) this.mnuRPtComprasCA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			 
    

    
					if (!a[1026]) this.mnuRPtVentas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1041]) this.mnuRPTVentasPorcMarca.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1042]) this.mnuRPTVentasPorcMarcaUtilid.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1043]) this.mnuRPTVentasMargUti.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1044]) this.mnuRPTVentasMaUXArt.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1045]) this.mnuRPTVentasAnInMEn.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1046]) this.mnuRPTVentasRevFact.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1047]) this.mnuRPTVentasXDia.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1048]) this.mnuRPTVentasNFactPForPago.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1049]) this.mnuRPTVentasGaraExt.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1050]) this.mnuRPTVentasXFecha.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1051]) this.mnuRPTVentasComDeVentas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1052]) this.mnuRPTVentasANuXLo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1053]) this.mnuRPTVentasMenXLo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1054]) this.mnuRPTVentasTarCredi.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1055]) this.mnuRPTVentasDetTc.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1056]) this.mnuRPTVentasEncuestas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1057]) this.mnuRPTVentasReCXCo.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1058]) this.mnuRPTVentasXano.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1059]) this.mnuRPTVentasXArtic.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1060]) this.mnuRPTVentasArtXMes.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//															if (!a[1151]) this.barSubItem12.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//																	if (!a[1026]) this.mnuRPTVentasComisionesConf.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//																	if (!a[1026]) this.mnuRPTVentasComisionesCali.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//																	if (!a[1026]) this.mnuRPTVentasComisionesDetGe.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
//																	if (!a[1026]) this.mnuRPTVentasComisionesConsulta.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

			
				if (!a[1231]) this.mnuReporVentasDFI.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;



			       if (!a[1144]) this.mnuRPTVentasPorcMargUtilidad.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			      if (!a[1174]) this.mnuReporVentasPDGA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			


			        
					if (!a[1027]) this.mnuRPtCREPOINT.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
							#region credipoint
			           	if (!a[1226]) this.mnuRPtCREPOINTFDG.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
				          if (!a[1228]) this. mnuRPtCREPOINRMD.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			             
							#endregion credipoint
					if (!a[1028]) this.mnuRPtDT.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
					if (!a[1029]) this.mnuRPtPeonyServicios.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			#endregion Reporteria

			bAbrePromocionesActivas = a[817];
			bAbreListaPersonalPrueba = a[818];
			bAbreListaClientesRevisados = a[819];
			bAbreListaClientesAtrasados = a[820];

			if (!a[821]) this.mnuMetricas.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[822]) this.mnuPCGEx.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[823]) this.mnuSTA.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1143]) this.mnuIngreso.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
			if (!a[1176]) this.mnuCPCobCM.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;			
		}

		/// <summary>
		/// Punto de entrada principal de la aplicación.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			//			try
			//			{
			Application.Run(new MenuLatinium());
			//			}
			//			catch(System.Exception ex)
			//			{
			//				MessageBox.Show(ex.Message, "Error Capturado en main");
			//			}
		}

		#region Variables
		public Acceso miAcceso = new Acceso();
		static public string stDirServidor = "";
		static public string stDirReportes = "";
		static public string stDirConta = "";
		static public string stDirFacturacion = "";
		static public string stDirNumero = "";
		static public string stDirRetenc = "";
		static public string stDirRoles = "";
		static public string stDirInicio = "";
		static public int iTipoDB = 1;
		static public int iNivel = 1;
		static public int IdUsuario = 1;
		static public string stUsuario = "Prueba1";
		static public int iRegistro = 1;
		static public string ImpresoraNombre;
		static public bool bFox=false;
		static public string sRuta="" ;
		static public bool bAuditoria = false;
		static public bool bAcademica = false;
		static public DateTime dtFechaSistema = DateTime.Today;
		static public int stIdDB = 0;
		static public bool bComisionQuincena = false;
		static public string sVersion = "";
		static public bool bConsignacionVenta = false;
		#endregion Variables

		private void Form1_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			stDirInicio = Application.StartupPath;
			if (stDirInicio.EndsWith("Debug")) stDirInicio = @"C:\Latinium";
			if (stDirInicio.EndsWith("Release")) stDirInicio = @"C:\Latinium";
			this.InicioSesion();
			//if (MenuLatinium.IdUsuario == 44) this.timer1.Start();
			//if (MenuLatinium.IdUsuario == 1) this.mnuVentasVentasCP.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
			Cursor = Cursors.Default;			
		}

		private void MnuCont_Plan_Ctas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			PlanCuentas miPlanCuenta =new PlanCuentas();
			miPlanCuenta.MdiParent=this;
			miPlanCuenta.Show();
			Cursor = Cursors.Default;
		}

		private void Mnu_EdicSeguridad_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Seguridad miSeguridad =  new Seguridad();
			miSeguridad.ShowDialog();
			miSeguridad.Dispose();
			cdsSeguridad.Clear();
			cdsSeguridad.Fill();
			AsignaMenu(cdsSeguridad.TableViews["Seguridad"], miAcceso.NivelAcceso);
			Cursor = Cursors.Default;
		}

		private void MnuContRepCont_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmReportesDeContabilidad Reporte =  new frmReportesDeContabilidad();	
			Reporte.MdiParent = this;
			Reporte.Show();

			frmInformesDeContabilidad IDC = new frmInformesDeContabilidad();
			IDC.MdiParent = this;
			IDC.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactArtic_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmArticulos Articulos = new frmArticulos();
			Articulos.MdiParent = this;
			Articulos.Show();
			Cursor = Cursors.Default;
		}

		private void mnuProdClientes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ProdPronostico miPronostico = new ProdPronostico();
			miPronostico.MdiParent = this;
			miPronostico.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactCompras_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			try
			{
				Cursor = Cursors.WaitCursor;
				frmCompra Compra = new frmCompra();
				Compra.MdiParent = this;
				Compra.Show();
				Cursor = Cursors.Default;

				//IngresaCompra(4);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Compra");
			}
		}

		private void IngresaCompra(int idTipo)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			//			Compra miCompra = new Compra(idTipo);
			//			miCompra.MdiParent = this;
			//			miCompra.Show();
			//			miCompra.Visibilidad();
			DateTime dtFin = DateTime.Now;
			TimeSpan ts = dtFin.Subtract(dtInicio);
			lblMensaje.EditValue = "Tiempo: " + ts.TotalSeconds.ToString("0.00");
			Cursor = Cursors.Default;
		}

		private bool EsPuntoVenta()
		{
			if (Funcion.bEjecutaSQL(cdsSeguridad, "Exec SeteoGExiste 'FPVTPVT'"))
			{
				Venta miVenta = new Venta();
				miVenta.MdiParent = this;
				miVenta.Show();
				return true;
			}
			return false;
		}

		private void MnuFactVentas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			try
			{
				//				if (EsPuntoVenta()) return;
				
				Cursor = Cursors.WaitCursor;
				frmVenta Ventas = new frmVenta();
				Ventas.MdiParent = this;
				Ventas.Show();
				Cursor = Cursors.Default;

				//				IngresaCompra(1);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Venta");
			}
		}

		private void MnuFactOrdComp_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			try
			{
				//IngresaCompra(2);
				//				frmOrdenDeCompra odc = new frmOrdenDeCompra();
				//				odc.MdiParent = this;
				//				odc.Show();

				frmOrdenesDeCompra OrdenDeCompra = new frmOrdenesDeCompra();
				OrdenDeCompra.MdiParent = this;
				OrdenDeCompra.Show();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Orden de compra");
			}
		}

		private void MnuFactCotizacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresaCompra(3);
		}

		private void MnuFactDevVentas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			frmDevolucionVenta miDevolucion = new frmDevolucionVenta();
			miDevolucion.MdiParent = this;
			miDevolucion.Show();
			DateTime dtFin = DateTime.Now;
			TimeSpan ts = dtFin.Subtract(dtInicio);
			lblMensaje.EditValue = "Tiempo: " + ts.TotalSeconds.ToString("0.00");
			Cursor = Cursors.Default;
			//			IngresaCompra(5);
		}

		private void MnuFactRepInventario_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			RepInventario miRepInventario = new RepInventario();
			miRepInventario.MdiParent = this;
			miRepInventario.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactDevCompras_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			IngresaCompra(6);
			frmDevolucionCompra miDevolucionProv = new frmDevolucionCompra();
			miDevolucionProv.MdiParent = this;
			miDevolucionProv.Show();
		}

		private void MnuFactRecTransferencia_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmIngresosDeBodega IngresosDeBodega = new frmIngresosDeBodega();
			IngresosDeBodega.MdiParent = this;
			IngresosDeBodega.Show();
		}

		private void MnuFactEnvTransferencia_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//IngresaCompra(8);
			frmEgresosDeBodega EgresoBodega = new frmEgresosDeBodega();
			EgresoBodega.MdiParent = this;
			EgresoBodega.Show();
		}

		private void MnuProducRecetas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Receta miReceta = new Receta();
			miReceta.MdiParent = this;
			miReceta.Show();
			Cursor = Cursors.Default;
		}

		private void MnuCont_Comparacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			bool bFormMarzo = false;
			string stSelect = "Select FormMarzo From Seteo";
			bFormMarzo = Funcion.bEscalarSQL(cdsSeguridad, stSelect, false);
			if (!bFormMarzo)
			{
				//				Egreso miEgreso = new Egreso();
				//				miEgreso.MdiParent = this;
				//				miEgreso.Show();
			}
			else
			{
				//				EgresoAT miEgreso = new EgresoAT();
				//				miEgreso.MdiParent = this;
				//				miEgreso.Show();
			}
			Cursor = Cursors.Default;
		}

		private void MnuArchSalir_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Close();
		}

		private void AbrePagos(int iTipo)
		{
			Cursor = Cursors.WaitCursor;
			string stSelect = "";

			if (iTipo == 1)
			{
				stSelect = "SELECT COUNT(*) FROM Compra "
					+ "WHERE (idTipoFactura = 1 Or idTipoFactura = 12) "
					+ "And ContadoCredito = 2";
			}
			else
			{
				stSelect = "SELECT COUNT(*) FROM Compra "
					+ "WHERE (idTipoFactura = 4 Or idTipoFactura = 11) "
					+ "And ContadoCredito = 2";
			}

			int iCuenta = 0;
			iCuenta = Funcion.iEscalarSQL(cdsSeguridad,stSelect);

			if (iCuenta == 0)
			{
				if (iTipo == 1)
					MessageBox.Show("No hay informacion para realizar Cobros");
				else
					MessageBox.Show("No hay informacion para realizar Pagos");

				Cursor = Cursors.Default;
				return;
			}

			//			Pagos miPago = new Pagos(iTipo);
			//			miPago.MdiParent = this;
			//			miPago.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactCtasCobrar_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//AbrePagos(1);
			frmCuentasPorCobrar miCuenta = new frmCuentasPorCobrar();
			miCuenta.MdiParent = this;
			miCuenta.Show();
		}

		private void MnuFactRepResumenes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			RepFacturacion miReporte =new RepFacturacion();
			miReporte.MdiParent = this;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactGraficos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			RepArticulos miReporte = new RepArticulos();
			miReporte.MdiParent = this;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void MnuContEstadoCta_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//			EstadoCuenta miEstado = new EstadoCuenta();
			//			miEstado.MdiParent = this;
			//			miEstado.Show();

			frmConciliacionBancaria Conciliacion = new frmConciliacionBancaria();
			Conciliacion.MdiParent = this;
			Conciliacion.Show();

			Cursor = Cursors.Default;
		}

		private void MnuFactClientes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Cliente miCliente = new Cliente(true);
			miCliente.MdiParent = this;
			miCliente.Show();

			//			frmClientesProveedores CP = new frmClientesProveedores(false);
			//			CP.MdiParent = this;
			//			CP.Show();
			//			Cursor = Cursors.Default;
		}

		private void MnuFactProvee_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Cliente miCliente = new Cliente(false);
			miCliente.MdiParent = this;
			miCliente.Show();		
			Cursor = Cursors.Default;
		}

		private void MnuFactGrupArticulos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ArtGrupo miGrupo = new ArtGrupo();
			miGrupo.MdiParent = this;
			miGrupo.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactProyectos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Proyecto miProyecto = new Proyecto();
			miProyecto.MdiParent = this;
			miProyecto.Show();
			Cursor = Cursors.Default;
		}

		private void MnuRolPersonal_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmNomina Nomina = new frmNomina();
			Nomina.MdiParent = this;
			Nomina.Show();			
			Cursor = Cursors.Default;
		}

		private void MnuProducActualiza_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
		}

		private void MnuProducReportes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			RepProduccion miReporte = new RepProduccion();
			miReporte.MdiParent = this;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void barButtonItem9_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ProductoMes miProducto = new ProductoMes();
			miProducto.MdiParent = this;
			miProducto.Show();
			Cursor = Cursors.Default;
		}

		private void MnuHerr_Resumen_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ActualizaSaldo miSaldo = new ActualizaSaldo();
			miSaldo.MdiParent = this;
			miSaldo.Show();
			Cursor = Cursors.Default;
		}

		private void MnuProducOrdProd_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;			
			frmGeneraOrdenesDeProduccion ODP = new frmGeneraOrdenesDeProduccion();
			ODP.MdiParent = this;
			ODP.Show();
			Cursor = Cursors.Default;
		}

		private void MnuProducProdTerm_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresaCompra(12);
		}

		private void mnuFactTransf_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmTransferencias Transferencias = new frmTransferencias();
			Transferencias.MdiParent = this;
			Transferencias.Show();
			Cursor = Cursors.Default;
		}

		private void mnuProdTransf_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//			CompraTransf miCopiaTransf = new CompraTransf();
			//			miCopiaTransf.MdiParent = this;
			//			miCopiaTransf.Show();		
			Cursor = Cursors.Default;
		}

		private void InicioSesion()
		{
			Inicio miInicio = new Inicio();
			if (miInicio.ShowDialog() != DialogResult.OK) 
			{
				Application.Exit();
				return;
			}

			miAcceso.NivelAcceso = miInicio.iNivel;
			iNivel = miAcceso.NivelAcceso;
			IdUsuario = miInicio.idUsuario;
			this.Text = miInicio.stNombreEmpresa.Trim();
			
			miInicio.Dispose();
			
			if (!ActivaSeguridad())
			{
				Application.Exit();
				return;
			}

			if (!Funcion.bEscalarSQL(cdsSeguridad, string.Format("Exec ValidaAccesoPorMaquina {0}", MenuLatinium.IdUsuario), true))
			{
				MessageBox.Show("No tiene Acceso al sistema en esta maquina con el nombre de usuario ingresado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
				Application.Exit();
				return;
			}

			string sSQL = string.Format("Exec ValidaEstadoContraseñaAcceso '{0}', '{1}', {2}",
				miInicio.txtNombre.Text.ToString(), miInicio.txtClave.Text.ToString(), MenuLatinium.IdUsuario);
			string sMensaje = Funcion.sEscalarSQL(cdsSeguridad, sSQL);

			if (sMensaje.Length > 0)
			{
				using (UsuarioClave Clave = new UsuarioClave())
				{				
					Clave.ShowDialog();
					Application.Exit();					
				}
			}
			
			Funcion.EjecutaSQL(cdsSeguridad, string.Format("Exec InicioActualizaBloqueoTransacciones '{0}'", miInicio.txtNombre.Text.ToString()));

			dtFechaSistema = Funcion.dtEscalarSQL(cdsSeguridad, "Select GETDATE()");

			ValidaComisionQuincena();

			if (bAbrePromocionesActivas)
			{
				string sSQLValPromocion = string.Format("Select COUNT(*) From DescripcionPromociones Where Estado = 1");
				if (Funcion.iEscalarSQL(cdsSeguridad, sSQLValPromocion, true) > 0)
				{
					frmVistaPromociones Vista = new frmVistaPromociones();
					Vista.MdiParent = this;
					Vista.Show();
				}
			}						
			
			if (bAbreListaPersonalPrueba)
			{
				frmNominaAvisoPersonalPrueba NAPP = new frmNominaAvisoPersonalPrueba();
				NAPP.MdiParent = this;
				NAPP.Show();
			}

			if (bAbreListaClientesRevisados)
			{
				frmCre_ListaClientesCreditoRevisados CLCR = new frmCre_ListaClientesCreditoRevisados();
				CLCR.MdiParent = this;
				CLCR.Show();
			}

			if (bAbreListaClientesAtrasados)
			{
				frmCre_ListaAbonosClientesAtrasados LACA = new frmCre_ListaAbonosClientesAtrasados();				
				LACA.Show();
			}

			if (Funcion.bEscalarSQL(cdsSeguridad, string.Format("Select Recordatorios From Usuario Where idUsuario = {0}", MenuLatinium.IdUsuario), true))
			{
				if (Funcion.iEscalarSQL(cdsSeguridad, "Select COUNT(*) From RecordatoriosCumplimientos Where Estado = 0 And DATEDIFF(DAY, GETDATE(), CONVERT(Date, Fecha)) <= 5") > 0)
				{
					frmListaCumplimientosPendientes LCP = new frmListaCumplimientosPendientes();
					LCP.ShowDialog();
				}
			}			

			MenuLatinium.dtFechaSistema = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeguridad);
			bConsignacionVenta = Funcion.bEscalarSQL(cdsSeguridad, "Select ConsignacionVenta From SeteoF", true);
			string stAudita = "Exec AuditaCrear 53, 6, 'Inicio'";
			Funcion.EjecutaSQL(cdsSeguridad, stAudita, true);			
		}

		private void ValidaComisionQuincena()
		{
			bComisionQuincena = Funcion.bEscalarSQL(cdsSeguridad, "Select Quincena From Seteo", true);

			if (bComisionQuincena) this.mnuRolesPagoComisiones.Caption = "Generar Quincena";
		}

		string CadenaConexion = "";
		public int idUsuarioInicio= 0;

		private void InicioSesionBase()
		{
			DBIngresoUsuario miInicio = new DBIngresoUsuario();
			if (miInicio.ShowDialog() != DialogResult.OK) 
			{
				Application.Exit();
				return;
			}

			miAcceso.NivelAcceso = miInicio.idGrupo;
			iNivel = miAcceso.NivelAcceso;
			miInicio.Dispose();
			CadenaConexion = miInicio.stConeccion;
			idUsuarioInicio = miInicio.vidUsuario;
			DBAccedeEmpresa miInicio1 = new DBAccedeEmpresa(miInicio.stConeccion, miInicio.vidUsuario);
			if (miInicio1.ShowDialog() != DialogResult.OK) 
			{
				Application.Exit();
				return;
			}
			this.Text = miInicio1.stNombreEmpresa.Trim();
			#region verificar seguridades timpo clave caducidad
			this.cdsSeguridad.Clear();
			Funcion miFuncion = new Funcion();
			cdsSeguridad.Schema.Connections[0].ConnectionString = CadenaConexion;
			string stProced = "exec UsuarioSeguridades "+idUsuarioInicio;
			if(	Funcion.iEscalarSQL(cdsSeguridad,stProced) == 1)
			{
				Cursor = Cursors.WaitCursor;
				DBCambioClave miUsuario = new DBCambioClave(idUsuarioInicio);
				if (miUsuario.ShowDialog() != DialogResult.OK) 
				{
					Application.Exit();
					Cursor = Cursors.Default;
					return;
				}
				this.cdsSeguridad.Clear();
				cdsSeguridad.Schema.Connections[0].ConnectionString = CadenaConexion;
				Funcion.EjecutaSQL(cdsSeguridad,"exec UsuarioSeguridades "+idUsuarioInicio+",1");
				Cursor = Cursors.Default;
			}
			#endregion verificar seguridades timpo clave caducidad

			if (!ActivaSeguridad())
			{
				Application.Exit();
				return;
			}
			#region Verificacion de numero de usuarios
			if (File.Exists(@"C:\Latinium\Academico.txt"))
			{
				bAcademica = true;
				MessageBox.Show("Activada version academica", "Información");
			}
			else if (!File.Exists(@"C:\Latinium\libAnexo.dll"))
			{
				try
				{
					string stExec = "Exec AuditaUsuarios";
					int iUsuarios = Funcion.iEscalarSQL(cdsSeguridad, stExec, true);
					if (iUsuarios < 0)
					{
						MessageBox.Show("Sistema Bloqueado el dia de Hoy por Limite de Usuarios superado. \nComuniquese con Infoelect para adquirir mas licencias", "Información");
						Registro miReg = new Registro();
						miReg.ShowDialog();
						Application.Exit();
						return;
					}
					stExec = "Exec AuditaFecha";
					int iFecha = Funcion.iEscalarSQL(cdsSeguridad, stExec, true) * -1;
					if (iFecha > 0)
					{
						string stErrorFecha = "";
						if (iFecha != 100) stErrorFecha = "Existe un retrazo en los pagos de " + iFecha.ToString() +" dias.";
						MessageBox.Show(stErrorFecha + " Activada version academica", "Información");
						bAcademica = true;
					}
				}
				catch//(System.Exception ex)
				{
					MessageBox.Show("Error en activacion", "Información");
				}
			}
			#endregion Verificacion de numero de usuarios
			string stAudita = "Exec AuditaCrear 53, 6, 'Inicio'";
			Funcion.EjecutaSQL(cdsSeguridad, stAudita, true);
			Funcion.EjecutaSQL(cdsSeguridad, "If Exists (SELECT * FROM sys.objects Where Name = 'AsientoBorraVacio') Exec AsientoBorraVacio");
		}

		private bool ActivaSeguridad()
		{
			// Activa de Acuerdo a Directorio de Seguridad
			this.cdsSeguridad.Clear();
			Funcion miFuncion = new Funcion();
			cdsSeguridad.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			
			try
			{
				#region Ingreso Normal
				try
				{
					cdsSeguridad.Clear();
					cdsSeguridad.Fill();
				}
				catch
				{
					#region Cambio de Contraseña
					//					MessageBox.Show(string.Format("Verificando posible cambio de contraseña",
					//						MenuLatinium.stDirServidor, MenuLatinium.stDirFacturacion), "Informacion");
					cdsSeguridad.Schema.Connections[0].ConnectionString = 
						cdsSeguridad.Schema.Connections[0].ConnectionString.Replace("Bl45o6$9", "infoelect2243");
					cdsSeguridad.Clear();
					cdsSeguridad.Fill();

					int iCuenta = Funcion.iEscalarSQL(cdsSeguridad, 
						"SELECT Count(*) FROM bases.sys.objects Where Name = 'UsuariosClave'");
					if (iCuenta == 1)
					{
						try
						{
							Funcion.EjecutaSQL(cdsSeguridad, "Exec Bases.dbo.UsuariosClave 'Bl45o6$9'");
							MessageBox.Show("Cambio de clave. Reinicie el programa", "Informacion",
								MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
						catch
						{
							MessageBox.Show("Ejecute el procedimiento UsuariosClave desde bases", "Informacion",
								MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
						return false;
					}
					#endregion Cambio de Contraseña
				}
							
				FileVersionInfo Info = FileVersionInfo.GetVersionInfo(@"C:\Latinium\Latinium.exe");				
				sVersion = "Version: " + Info.FileVersion;				
				this.Text += " - " + MenuLatinium.stUsuario + " - " + sVersion;
				AsignaMenu(cdsSeguridad.TableViews["Seguridad"], miAcceso.NivelAcceso);
				return true;
				#endregion Ingreso Normal
			}
			catch
			{
				#region Error no hay servidor base o usuario
				MessageBox.Show(string.Format("No pudo ingresar a Servidor {0} Base de datos {1} el usuario {2}", 
					MenuLatinium.stDirServidor, MenuLatinium.stDirFacturacion, MenuLatinium.stUsuario), "Informacion");

				return false;
				#endregion Error no hay servidor base o usuario

				#region Bloqueado
				#region Variables globales
				bool VerificaServidor =  false;
				string stSelect = "";
								
				C1.Data.SchemaObjects.Connection miConeccion = cdsSeguridad.Schema.Connections[0];
				#endregion Variables globales

				try
				{
					#region Verificacion de ingreso a sevidor con autenticacion de Windows
					cdsSeguridad.Schema.Connections[0].ConnectionString = 
						@"Integrated Security=SSPI;Persist "
						+ "Security Info=False;" 
						+ "Pooling=False; Connection Timeout=30;"
						+	"Initial Catalog=Master"
						+ ";Data Source=" + MenuLatinium.stDirServidor.Trim();
					miConeccion.Open();
					VerificaServidor =  true;
					#endregion Fin de windows 
				}
				catch // No puede abrir con usuario de Windows
				{
					#region Usuario Infoelect
					MessageBox.Show(string.Format("No pudo ingresar a Servidor {0} el usuario Windows",
						MenuLatinium.stDirServidor), "Informacion");
					cdsSeguridad.Schema.Connections[0].ConnectionString = @"Password=Bl45o6$9;"
						+ "Persist Security Info=True;User ID=infoelect"
						+ ";Initial Catalog=Master"
						+ ";Data Source=" + MenuLatinium.stDirServidor.Trim();
					miConeccion = cdsSeguridad.Schema.Connections[0];
					try
					{
						miConeccion.Close();
						miConeccion.Open();
						VerificaServidor =  true;
					}
					catch // No puede abrir con usuario infoelect
					{
						MessageBox.Show(string.Format("No se puede abrir la Instancia {0} con usuario Infoelect ni windows",
							MenuLatinium.stDirServidor), "Informacion");
					}
					#endregion Usuario Infoelect
				}

				#region Ingresa a servidor
				if (VerificaServidor)
				{
					#region No existe base de datos
					stSelect = string.Format("SELECT Count(*) FROM master.dbo.sysdatabases Where name = '{0}'",
						MenuLatinium.stDirFacturacion);
					int iCuentaBases = Funcion.iEscalarSQL(cdsSeguridad, stSelect);
					if (iCuentaBases == 0)
					{
						MessageBox.Show(string.Format("No existe base de datos {0} Adjuntada", 
							MenuLatinium.stDirFacturacion), "Informacion");
					}
					#endregion No existe base de datos
					#region No existe usuario
					stSelect = string.Format("SELECT Count(*) FROM Master.Dbo.syslogins Where LoginName = '{0}'",
						MenuLatinium.stUsuario);
					iCuentaBases = Funcion.iEscalarSQL(cdsSeguridad, stSelect);
					if (iCuentaBases == 0)
					{
						stSelect = string.Format("Create Login {0} With Password = '{1}'; Exec sp_addsrvrolemember '{0}', 'sysadmin';"
							+ "Create User {0} For Login {0}",
							MenuLatinium.stUsuario, "Bl45o6$9");
						Funcion.EjecutaSQL(cdsSeguridad, stSelect);
						MessageBox.Show(string.Format("Usuario {0} creado", 
							MenuLatinium.stUsuario), "Informacion");
					}
					#endregion No existe usuario
					miConeccion.Close();
				}
				#endregion Ingresa a servidor

				Empresa miEmpresa = new Empresa();
				miEmpresa.ShowDialog();
				return false;				
				#endregion Bloqueado
			}
		}

		private void MnuArchIniSes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			if (System.IO.File.Exists(@"C:\Latinium\LibAnexos1.dll"))
			{
				DBAccedeEmpresa miAbrir = new DBAccedeEmpresa(CadenaConexion,idUsuarioInicio);
				if (miAbrir.ShowDialog() != DialogResult.OK) 
				{
					return;
				}
				this.Text = miAbrir.stNombreEmpresa.Trim();
				miAbrir.Dispose();
				string stAudita = "Exec AuditaCrear 53, 7, 'Salida Latinium'";
				Funcion.EjecutaSQL(cdsSeguridad, stAudita, true);
				ActivaSeguridad();
			}
			else
			{
				Cursor = Cursors.WaitCursor;
				stDirInicio = Application.StartupPath;
				if (stDirInicio.EndsWith("Debug")) stDirInicio = @"C:\Latinium";
				if (stDirInicio.EndsWith("Release")) stDirInicio = @"C:\Latinium";
				this.InicioSesion();
				Cursor = Cursors.Default;

				//				AbrirEmpresa miAbrir = new AbrirEmpresa();
				//				if (miAbrir.ShowDialog() != DialogResult.OK) 
				//				{
				//					return;
				//				}
				//				this.Text = miAbrir.stNombreEmpresa.Trim();
				//				miAbrir.Dispose();
				//				string stAudita = "Exec AuditaCrear 53, 7, 'Salida Latinium'";
				//				Funcion.EjecutaSQL(cdsSeguridad, stAudita, true);
				//				ActivaSeguridad();
			}
		}

		private void MnuArchNuevo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			if (!this.MnuArchNuevo.Enabled)
			{
				MessageBox.Show("No tiene acceso a la creación de nuevas empresas",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			//			Cursor = Cursors.WaitCursor;
			//			NuevaEmpresa miEmpresa = new NuevaEmpresa();			
			//			miEmpresa.MdiParent = this;
			//			miEmpresa.Show();

			frmCreaEmpresas Empresa = new frmCreaEmpresas();
			Empresa.MdiParent = this;
			Empresa.Show();
			Cursor = Cursors.Default;
		}

		private void MnuContTransE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			bool bFormMarzo = false;
			string stSelect = "Select FormMarzo From Seteo";
			bFormMarzo = Funcion.bEscalarSQL(cdsSeguridad, stSelect, false);
			if (!bFormMarzo)
			{
				//				Egreso miEgreso = new Egreso();
				//				miEgreso.MdiParent = this;
				//				miEgreso.Show();
			}
			else
			{
				//				EgresoAT miEgreso = new EgresoAT();
				//				miEgreso.MdiParent = this;
				//				miEgreso.Show();
			}
			Cursor = Cursors.Default;
		}

		private void MnuContTransI_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//			IngresoAT miIngreso = new IngresoAT();
			//			frmRetVenta miIngreso = new frmRetVenta();
			//			miIngreso.MdiParent = this;
			//			miIngreso.Show();		
			Cursor = Cursors.Default;
		}

		private void MnuHerr_Actu_Vers_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ActualizaVersion miActualizacion = new ActualizaVersion();
			miActualizacion.MdiParent = this;
			miActualizacion.Show();
		}

		private void MnuFactCtasPagar_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//AbrePagos(4);
			frmCuentasPorPagar miCuenta = new frmCuentasPorPagar();
			miCuenta.MdiParent = this;
			miCuenta.Show();
		}

		private void bbiUsua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{

			if (System.IO.File.Exists(@"C:\Latinium\LibAnexos1.dll"))
			{
				Cursor = Cursors.WaitCursor;
				DBUsuario miUsuario = new DBUsuario();
				miUsuario.MdiParent = this;
				miUsuario.Show();
				Cursor = Cursors.Default;
			}
			else
			{
				Cursor = Cursors.WaitCursor;
				//				Usuario miUsuario = new Usuario();
				frmUsuarios miUsuario = new frmUsuarios();
				//				frmGuardaUsuarios miUsuario = new frmGuardaUsuarios();
				miUsuario.MdiParent = this;
				miUsuario.Show();
				Cursor = Cursors.Default;
			}
		}

		private void mnuArchPropEmpresa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			if (System.IO.File.Exists(@"C:\Latinium\LibAnexos1.dll"))
			{
				DBEmpresaBD miEmpresa = new DBEmpresaBD();
				miEmpresa.MdiParent = this;
				miEmpresa.Show();
			}
			else
			{
				Empresa miEmpresa = new Empresa();
				miEmpresa.MdiParent = this;
				miEmpresa.Show();
			}
			Cursor = Cursors.Default;
		}

		private void mnuArchPropGen_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Seteo miSeteo = new Seteo();
			miSeteo.MdiParent = this;
			miSeteo.Show();
			Cursor = Cursors.Default;
		}

		private void barNuevo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			NuevaEmpresa miEmpresa = new NuevaEmpresa();
			miEmpresa.ShowDialog();
			miEmpresa.Dispose();
			Cursor = Cursors.Default;
		}

		private void MnuEdicCortar_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{

		}

		private void MnuHerr_Camb_Contr_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			if (System.IO.File.Exists(@"C:\Latinium\LibAnexos1.dll"))
			{
				Cursor = Cursors.WaitCursor;
				DBCambioClave miUsuario = new DBCambioClave(idUsuarioInicio);
				miUsuario.MdiParent = this;
				miUsuario.Show();
				Cursor = Cursors.Default;
			}
			else
			{
				Cursor = Cursors.WaitCursor;
				UsuarioClave miUsuario = new UsuarioClave();
				//				miUsuario.MdiParent = this;
				miUsuario.ShowDialog();
				Cursor = Cursors.Default;
			}
		}

		private void MnuArchGuardar_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			GrabarBase miGrabacion = new GrabarBase();
			miGrabacion.MdiParent = this;
			miGrabacion.Show();
		}

		private void MnuContTransR_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//			ResumenRetencion miRetencion = new ResumenRetencion();
			//			miRetencion.MdiParent = this;
			//			miRetencion.Show();

			frmReporteRetenciones Reporte = new frmReporteRetenciones();
			Reporte.MdiParent = this;
			Reporte.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactActFijos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ActivoFijo miActivo = new ActivoFijo();
			miActivo.MdiParent = this;
			miActivo.Show();
			Cursor = Cursors.Default;
		}

		private void mnuHerrGenSQL_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			GeneraSQL miGenera = new GeneraSQL();
			miGenera.MdiParent = this;
			miGenera.Show();
			Cursor = Cursors.Default;
		}

		private void MnuProducConsReal_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresaCompra(10);
		}

		private void MnuFactLot_Ctas_Cob_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			PagoLotes miPagoLotes = new PagoLotes(1);
			miPagoLotes.MdiParent = this;
			miPagoLotes.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactLot_Ctas_Pag_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			PagoLotes miPagoLotes = new PagoLotes(4);
			miPagoLotes.MdiParent = this;
			miPagoLotes.Show();		
			Cursor = Cursors.Default;
		}

		private void barRolRubro_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			RolConcepto miConcepto = new RolConcepto();
			miConcepto.MdiParent = this;
			miConcepto.Show();
			Cursor = Cursors.Default;
		}

		private void MnuRolGenRol_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//			RolGenera miGenera = new RolGenera();
			//			miGenera.MdiParent = this;
			//			miGenera.Show();
			frmRolesDePago Roles = new frmRolesDePago();
			Roles.MdiParent = this;
			Roles.Show();
			Cursor = Cursors.Default;
		}

		private void MnuRolVerif_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmRolPrestamo Prestamo = new frmRolPrestamo();
			Prestamo.MdiParent = this;
			Prestamo.Show();
			Cursor = Cursors.Default;
		}

		private void MnuHerr_Auditoría_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Auditoria miAuditoria = new Auditoria();
			miAuditoria.MdiParent = this;
			miAuditoria.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactImportacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//			Importacion miImportacion = new Importacion();
			//			miImportacion.MdiParent = this;
			//			miImportacion.Show();

			frmImportacion Importacion = new frmImportacion();
			Importacion.MdiParent = this;
			Importacion.Show();
			Cursor = Cursors.Default;
		}

		private void mnuHerrReporte_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ReporteLista miReporte = new ReporteLista();
			miReporte.MdiParent = this;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void mnuContCompara_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Comparacion miCompara = new Comparacion();
			miCompara.MdiParent = this;
			miCompara.Show();
			Cursor = Cursors.Default;
		}

		private void mnuContAsientoPeriodo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			AsientoPeriodo miPeriodo = new AsientoPeriodo();
			miPeriodo.MdiParent = this;
			miPeriodo.Show();
			Cursor = Cursors.Default;
		}

		private void MnuRolCargos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			RolCargo miCargo = new RolCargo();
			miCargo.MdiParent = this;
			miCargo.Show();
			Cursor = Cursors.Default;
		}

		private void MnuRolDpto_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			RolDepartamento miDepartamento = new RolDepartamento();
			miDepartamento.MdiParent = this;
			miDepartamento.Show();
			Cursor = Cursors.Default;
		}

		private void MnuRolRepRol_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			RepRoles miRol = new RepRoles();
			miRol.MdiParent = this;
			miRol.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactLot_Transf_Art_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			TransferenciaArticulos miTransf = new TransferenciaArticulos();
			miTransf.MdiParent = this;
			miTransf.Show();
			Cursor = Cursors.Default;
		}

		private void MnuArchConfImp_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ImpresoraPredefinida miImpresora = new ImpresoraPredefinida();
			miImpresora.MdiParent = this;
			miImpresora.Show();
			Cursor = Cursors.Default;
		}

		private void MnuEdicDesh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//this.ActiveMdiChild;//.barraDato1.barraDato1.PosUltima();
		}

		private void MnuFactLoteAnalisis_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ArticuloAnalisis miAnalisis = new ArticuloAnalisis();
			miAnalisis.MdiParent = this;
			miAnalisis.Show();
		}

		private void MnuContPresupuesto_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPresupuestoAnual PA = new frmPresupuestoAnual();
			PA.MdiParent = this;
			PA.Show();
		}

		private void mnuContLoteDepura_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Depuracion miDepura = new Depuracion();
			miDepura.MdiParent = this;
			miDepura.Show();
		}

		private void MnuRolAsigCtas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RolRenta miImpRenta = new RolRenta();
			miImpRenta.MdiParent = this;
			miImpRenta.Show();
		}

		//		private void mnuFactLotOrdenMov_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		//		{
		//			OrdenMovilizacion miOrdenMov = new OrdenMovilizacion();
		//			miOrdenMov.MdiParent = this;
		//			miOrdenMov.Show();
		//		}

		private void MnuHerr_Sal_Ini_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			SaldosIniciales miSaldo = new SaldosIniciales();
			miSaldo.MdiParent = this;
			miSaldo.Show();
		}

		private void barButtonItem5_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			System.Diagnostics.Process.Start("c:\\Latinium\\Latinium.chm");
		}

		private void MnuAyuda_Ayuda_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			System.Diagnostics.Process.Start("c:\\Latinium\\Latinium.chm");
		}

		private void mnuFactLotAgencia_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Reservacion miReserva = new Reservacion();
			miReserva.MdiParent = this;
			miReserva.Show();
		}

		private void MnuHerr_Copi_Arch_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			CopiarArchivos miCopia = new CopiarArchivos();
			miCopia.MdiParent = this;
			miCopia.Show();
		}

		private void MnuHerr_Camb_Ctes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			CambioConstantes miCambio = new CambioConstantes();
			miCambio.MdiParent = this;
			miCambio.Show();
		}

		private void MnuFactLot_Depositos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			PagoLoteDia miPagoDia = new PagoLoteDia();
			miPagoDia.MdiParent = this;
			miPagoDia.Show();
		}


		private void MnuFactLot_Facturacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			CompraLotes miFactLote = new CompraLotes();
			miFactLote.MdiParent = this;
			miFactLote.Show();
		}

		private void MnuFactPedClientes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			try
			{				
				frmPedidos Pedidos = new frmPedidos();
				Pedidos.MdiParent = this;
				Pedidos.Show();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error del Sistema al Abrir Pedidos");
			}
		}

		private void MnuFactNotaCV_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmNotaDeDebitoCliente NotaVenta = new frmNotaDeDebitoCliente();
			NotaVenta.MdiParent = this;
			NotaVenta.Show();
		}

		private void MnuFactNotaCC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmNotaDeDebitoProveedor NotaCompra = new frmNotaDeDebitoProveedor();
			NotaCompra.MdiParent = this;
			NotaCompra.Show();
		}

		private void MnuVentana_Casc_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			this.LayoutMdi(MdiLayout.Cascade);
		}

		private void MnuVentana_Org_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			this.LayoutMdi(MdiLayout.TileHorizontal);
		}

		private void MnuVentana_Cicl_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			this.LayoutMdi(MdiLayout.TileVertical);
		}

		private bool bTransparente = false;
		private void MnuVentTrans_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			bTransparente = !bTransparente;
		}

		private void mnuVentOpac_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			this.Opacity = this.Opacity - .1;
			if (Opacity < .7) Opacity = 1;
		}

		private void MnuFactPedPend(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			string stSelect = "Select Count(*) From SeteoG Where Codigo = 'FPEDC'";
			if (Funcion.iEscalarSQL(cdsSeguridad, stSelect, true) == 0)
			{
				CompraPedido miCompra = new CompraPedido();
				miCompra.MdiParent = this;
				miCompra.Show();
			}
			else
			{
				CompraPedido2 miCompra = new CompraPedido2();
				miCompra.MdiParent = this;
				miCompra.Show();
			}
		}

		private void mnuArchImpOmniLife_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			OmniLife miOmni = new OmniLife();
			miOmni.MdiParent = this;
			miOmni.Show();
		}

		private void mnuContTransAnexoTr_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmDetalleATS ATS = new frmDetalleATS();
			ATS.MdiParent = this;
			ATS.Show();
			//			AnexoTransaccional miAnexo = new AnexoTransaccional();
			//			miAnexo.MdiParent = this;
			//			miAnexo.Show();
			Cursor = Cursors.Default;
		}

		private void mnuFactLotesPago_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			PagoLoteDia miPagoDia = new PagoLoteDia(1);
			miPagoDia.MdiParent = this;
			miPagoDia.Show();
		}

		private void MnuFactReservClientes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//IngresaCompra(11);
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			ReservaCliente miReserva = new ReservaCliente();
			miReserva.MdiParent = this;
			miReserva.Show();
			DateTime dtFin = DateTime.Now;
			TimeSpan ts = dtFin.Subtract(dtInicio);
			lblMensaje.EditValue = "Tiempo: " + ts.TotalSeconds.ToString("0.00");
			Cursor = Cursors.Default;
		}

		private void MnuFactReservProveedores_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresaCompra(26);
		}

		private void MnuFactPedEmpleados_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresaCompra(29);		
		}

		private void mnuFactLoteRetazo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Retazo miRetazo = new Retazo();
			miRetazo.MdiParent = this;
			miRetazo.Show();
		}

		private void mnuContLotAsiComprob_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			AsientoGrupoComp miAsiGrupo = new AsientoGrupoComp();
			miAsiGrupo.MdiParent = this;
			miAsiGrupo.Show();
		}

		private void MnuFactResPedProv_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresaCompra(14);
		}

		private void mnuContTransFormulario_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Formulario103 miFormulario = new Formulario103();
			miFormulario.MdiParent = this;
			miFormulario.Show();
		}

		private void MenuLatinium_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Salir del Sistema?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				e.Cancel = true;
				return;
			}
			string stAudita = "Exec AuditaCrear 53, 7, 'Salida Latinium'";
			Funcion.EjecutaSQL(cdsSeguridad, stAudita, true);
		}

		private void mnuFactLotTomaF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			TomaFisica miToma = new TomaFisica();
			miToma.MdiParent = this;
			miToma.Show();
		}

		private void mnuFacLotCxC30_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Pagos_30_60 miPago = new Pagos_30_60();
			miPago.MdiParent = this;
			miPago.Show();
		}

		private void mnuAyuLiberar_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			long lMem1 = GC.GetTotalMemory(false);
			long lMem2 = GC.GetTotalMemory(true);
			GC.Collect();
			long lMem3 = GC.GetTotalMemory(false);
			long lMem4 = GC.GetTotalMemory(true);
			MessageBox.Show("Original: " + lMem1.ToString("#,##0")
				+ " Posterior: " + lMem3.ToString("#,##0"));
			MessageBox.Show("Total Original: " + lMem2.ToString("#,##0")
				+ " Posterior: " + lMem4.ToString("#,##0"));
		}

		private void MnuEdic_Cambiar_Reporte_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ReporteLista miReporte = new ReporteLista();
			miReporte.MdiParent = this;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void mnuArchImpTelcosan_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			SuperMas miTelcoSan = new SuperMas();
			miTelcoSan.MdiParent = this;
			miTelcoSan.Show();
		}

		private void MnuLotTransito_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			try
			{
				IngresaCompra(31);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Importacion en Transito");
			}		
		}

		private void MnuRolBloqRol_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RolBloqueo miRol = new RolBloqueo();
			miRol.MdiParent = this;
			miRol.Show();
		}

		private void mnuRolesLiquidacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RolLiquidacion miLiquidacion = new RolLiquidacion();
			miLiquidacion.MdiParent = this;
			miLiquidacion.Show();
		}

		private int iTiempo = 0;
		string[] ListaSkin = {"Caramel", "Money Twins", "Lilian", "Blue", "Imaginary", "The Asphalt World", "Black"};
		bool bProcesandoBasura = false;

		static public AlarmaMostrar _AlarmaMostrar;
		public AlarmaMostrar GetAlarmaMostrar
		{
			get
			{
				if (_AlarmaMostrar == null)
					_AlarmaMostrar = new AlarmaMostrar();
				
				if (_AlarmaMostrar.WindowState == System.Windows.Forms.FormWindowState.Minimized)
					_AlarmaMostrar.WindowState = System.Windows.Forms.FormWindowState.Normal;

				return _AlarmaMostrar;
			}
		} 

		private int iTiempoAlarma = 0;
		private int TiempoLimite = 0;//6;
		private void timer_Tick(object sender, System.EventArgs e)
		{
			if (bProcesandoBasura) return;
			#region Alarmas
			iTiempoAlarma ++;
			if (iTiempoAlarma >= TiempoLimite * 6 && TiempoLimite > 0)
			{
				try
				{
					iTiempoAlarma = 0;
					string stExec = "If Exists (SELECT * FROM sys.objects Where Name = 'AlarmaActivar') Exec AlarmaActivar '', 1";
					int iCuenta = Funcion.iEscalarSQL(cdsSeguridad, stExec);
					if (iCuenta == 0)
					{
						return;
					}
					else
						GetAlarmaMostrar.Show();
				}
				catch
				{
				}
			}
			#endregion alarmas

			if (Funcion.bAyudaRemota)
			{
				Form currentForm = this.ActiveMdiChild;
				if (currentForm != null)
					Funcion.CreaFormulario(currentForm);
			}
			iTiempo ++;
			if (iTiempo >= 60) // 60
			{
				bProcesandoBasura = true;
				DateTime dtIni = DateTime.Now;
				lblMensaje.EditValue = "T. Recol: 0.00";
				if (barSkin.EditValue.ToString() == "" || barSkin.EditValue.ToString() == "Rotativo")
				{
					Random mir = new Random();
					skinPrincipal.LookAndFeel.SetSkinStyle(ListaSkin[mir.Next(7)]);
				}
				iTiempo = 0;
				GC.Collect();
				TimeSpan ts = DateTime.Now.Subtract(dtIni);
				lblMensaje.EditValue = string.Format("T. Recol: {0}", ts.TotalSeconds.ToString("0.00"));
				bProcesandoBasura = false;
			}		
		}

		private void btHerrTester_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Tester miTester = new Tester();
			miTester.MdiParent = this;
			miTester.Show();
		}

		private void mneArchImpInterActive_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Interactive miInterActive = new Interactive();
			miInterActive.MdiParent = this;
			miInterActive.Show();
		}

		private void mnuFactLoteTarjeta_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Tarjeta miTarjeta = new Tarjeta();
			miTarjeta.MdiParent = this;
			miTarjeta.Show();
		}

		private void mnuFacLotSerie_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			CompraSerie miSerie = new CompraSerie();
			//			miSerie.MdiParent = this;
			//			miSerie.Show();
		}

		private void mnuRolAnexo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RolAnexo miRol = new RolAnexo();
			miRol.MdiParent = this;
			miRol.Show();
		}

		private void mneArchImpFuneraria_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			YellowPepper miYellow = new YellowPepper();
			miYellow.MdiParent = this;
			miYellow.Show();
		}

		private void mnuFacLotNichos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresaCompra(99);
		}

		private void MnuFactLot_Autoriza_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			PagosAutoriza miAutorizacion = new PagosAutoriza();
			miAutorizacion.MdiParent = this;
			miAutorizacion.Show();
		}

		private void MnuCont_Ind_Financieros_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IndiceFinanciero miIndice = new IndiceFinanciero();
			miIndice.MdiParent = this;
			miIndice.Show();
		}

		private void mneArchImpClaseEcuador_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ClaseEcaudor  miClaseEcaudor = new ClaseEcaudor();
			miClaseEcaudor.MdiParent = this;
			miClaseEcaudor.Show();
		}

		private void MnuAyudaAcercaCIAD_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			AcercaLatinium miAcercaDe = new AcercaLatinium();
			miAcercaDe.MdiParent = this;
			miAcercaDe.Show();
		}

		private void mnuProdLiquida_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			LiquidaOP miLiquida = new LiquidaOP();
			miLiquida.MdiParent = this;
			miLiquida.Show();
		}

		private void mneArchImpFidelius_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Fidelius  miFidelius = new Fidelius();
			miFidelius.MdiParent = this;
			miFidelius.Show();
		}

		private void mnuArchImpFresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			FreshLogistic miFresh = new FreshLogistic();
			miFresh.MdiParent = this.MdiParent;
			miFresh.Show();
		}

		private void mnuArchImpCash_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cash miCash = new Cash();
			miCash.MdiParent = this.MdiParent;
			miCash.Show();
		}

		private void MnuAyudaRegist_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Registro miRegistro = new Registro();
			miRegistro.MdiParent = this;
			miRegistro.Show();
		}

		private void mnuRolUtilidad_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RolUtilidad miRol = new RolUtilidad();
			miRol.MdiParent = this;
			miRol.Show();
		}

		private void MnuArchPropNumE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			CompraNumeroSri miNumero = new CompraNumeroSri();
			miNumero.MdiParent = this;
			miNumero.Show();
		}

		private void btProdCosto_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ProduccionCostos miProd = new ProduccionCostos();
			miProd.MdiParent = this;
			miProd.Show();
		}

		private void mnuArchImpGuiaGruas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			GuiaRemisionG  miGuiaRemisionG = new GuiaRemisionG();
			miGuiaRemisionG.MdiParent = this;
			miGuiaRemisionG.Show();		
		}

		private void mnuArchImpGuiaVenta_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			GuiaRemisionFactC miGuiaRemisionFactVenta = new GuiaRemisionFactC();
			miGuiaRemisionFactVenta.MdiParent = this;
			miGuiaRemisionFactVenta.Show();
		}

		private void mnuArchImpGruasCompra_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			GuiaRemisionFact   miGuiaRemisionFact = new GuiaRemisionFact();
			miGuiaRemisionFact.MdiParent = this;
			miGuiaRemisionFact.Show();		
		}

		private void mnuProdGeneraC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ProduccionGeneraConsumo miGeneraProdC = new ProduccionGeneraConsumo();
			miGeneraProdC.MdiParent = this;
			miGeneraProdC.Show();
		}

		private void mnuHerrSegFiltro_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ArticuloFiltro miArt = new ArticuloFiltro();
			miArt.MdiParent = this;
			miArt.Show();
		}

		private void mnuProdAgro_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			AgroLabor miAgro = new AgroLabor();
			miAgro.MdiParent = this;
			miAgro.Show();
		}

		private void mbuHerrSegTransf_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			SegTranforma miSeg = new SegTranforma();
			miSeg.MdiParent = this;
			miSeg.Show();
		}

		private void mnuHerrSegFiltroSuc_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			FiltroSucursal miFiltroS = new FiltroSucursal();
			miFiltroS.MdiParent = this;
			miFiltroS.Show();
		}

		private void mnuArchImpQBI_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			QueBuenaIdea miIdea = new QueBuenaIdea();
			miIdea.MdiParent = this;
			miIdea.Show();
		}

		private void btRequisicion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			try
			{
				IngresaCompra(36);		
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Requisición");
			}
		}

		private void mnuHerrSegApReq_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			AprobacionReq miAproReq = new AprobacionReq();
			miAproReq.MdiParent = this;
			miAproReq.Show();
		}

		private void MnuFactResPedComision_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Comision miComision = new Comision();
			miComision.MdiParent = this;
			miComision.Show();
		}

		private void mnuFactRepComp_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RepCompInv miRep = new RepCompInv();
			miRep.MdiParent = this;
			miRep.Show();
		}

		private void mnuHerrCredito_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RepCompCredito miRep = new RepCompCredito();
			miRep.MdiParent = this;
			miRep.Show();
		}

		private void btConfig_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			SeteoG miSeteoG = new SeteoG();
			miSeteoG.MdiParent = this;
			miSeteoG.Show();
		}

		private void mnuHerrBusqueda_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			FrmBusqueda miBusqueda = new FrmBusqueda();
			miBusqueda.MdiParent = this;
			miBusqueda.Show();
		}

		private void mnuFacLotCierre_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			CierreCaja miCierre = new CierreCaja();
			miCierre.MdiParent = this;
			miCierre.Show();
		}

		private void mnuHerrSegReqLis_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RequisicionListado miBusqueda = new RequisicionListado();
			miBusqueda.MdiParent = this;
			miBusqueda.Show();
		}

		private void mnuHerrSegReqEst_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RequisicionEstado miBusqueda= new RequisicionEstado();
			miBusqueda.MdiParent = this;
			miBusqueda.Show();
		}

		private void mnuArchImpCashE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			CashManagementEnvio miBusqueda= new CashManagementEnvio();
			miBusqueda.MdiParent = this;
			miBusqueda.Show();
		}

		private void mnuArchImpCerv_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cervecera miBusqueda = new Cervecera();
			miBusqueda.MdiParent = this;
			miBusqueda.Show();		
		}

		private void MnuContFlujo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCajaChica miFlujo = new frmCajaChica();			
			//Flujo miFlujo = new Flujo();
			miFlujo.MdiParent = this;
			miFlujo.Show();
		}

		private void MnuRolIncSuel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RolSueldoIncr miSueldo = new RolSueldoIncr();
			miSueldo.MdiParent = this;
			miSueldo.Show();
		}

		private void mnuArchImpResCDI_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ResumenCostosDI miCostosDI = new ResumenCostosDI();
			miCostosDI.MdiParent = this;
			miCostosDI.Show();
		}

		private void mnuArchImpCostoI_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IntCostoIndirecto miCostosI = new IntCostoIndirecto();
			miCostosI.MdiParent = this;
			miCostosI.Show();
		}

		private void mnuHerrContrato_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			
		}

		private void mnuRolesDistribucion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RolGeneraDistribucion rolDist = new RolGeneraDistribucion();
			rolDist.MdiParent = this;
			rolDist.Show();
		}				

		private void mnuFacLotTransfTrans_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			CompraCopia cc = new CompraCopia();
			cc.MdiParent = this;
			cc.Show();
		}

		private void mnuHerrContratosIngreso_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresoContratos ingC = new IngresoContratos();
			ingC.MdiParent = this;
			ingC.Show();
		}

		private void mnuHerrContratosGeneraFact_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngersoContratosFacturas ingF = new IngersoContratosFacturas();
			ingF.MdiParent = this;
			ingF.Show();
		}

		private void mnuHerrContratosAdemdums_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresoContratosAdemdums ad = new IngresoContratosAdemdums();
			ad.MdiParent =this;
			ad.Show();
		}

		private void mnuHerrContratosAlarma_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ContratosAlarma ca = new ContratosAlarma();
			ca.MdiParent = this;
			ca.Show();
		}

		private void mnuFacLotCopiaArt_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			CompraInsertaArticulo cia = new CompraInsertaArticulo();
			cia.MdiParent = this;
			cia.Show();
		}

		private void btSeprofin_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Seprofin miSeprofin = new Seprofin();
			miSeprofin.MdiParent = this;
			miSeprofin.Show();
		}

		private void btControlRemoto_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Funcion.bAyudaRemota = !Funcion.bAyudaRemota;
			if (Funcion.bAyudaRemota)
				btControlRemoto.Caption = "Conectado";
			else
				btControlRemoto.Caption = "Ayuda Remota";
		}

		private void mnuFacLotFactReembol_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			GeneraGastoVenta ad = new GeneraGastoVenta();
			ad.MdiParent =this;
			ad.Show();
		}

		private void btCuentaObligatoria_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			CuentaObliga miCuenta = new CuentaObliga();
			miCuenta.MdiParent = this;
			miCuenta.Show();
		}

		private void mnuRolesAdicional_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RolAdicional miAdicional = new RolAdicional();
			miAdicional.MdiParent = this;
			miAdicional.Show();
		}

		private void MnuFactResCruzarFactReserva_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ReservFact miCruce = new ReservFact();
			miCruce.MdiParent = this;
			miCruce.Show();
		}

		private void btMarceloGeneralVentas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			MarceloGeneralVentas miMarceloGeneralVentas = new MarceloGeneralVentas();
			//			miMarceloGeneralVentas.MdiParent = this;
			//			miMarceloGeneralVentas.Show();
		}

		private void mnuRolesAdicionalFormulas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RolAdicionalformulas miAdicionalF = new RolAdicionalformulas();
			miAdicionalF.MdiParent = this;
			miAdicionalF.Show();
		}

		private void MnuFactResOTrabajo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresaCompra(39); 
		}

		private void MnuFactResConclusion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresaCompra(40); 
		}

		private void MnuFactResOTrabajoServicio_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			IngresaCompra(41); 
		}

		private void mnuHerrPresupCIngreso_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			PresupuestoCompras miPresupC = new PresupuestoCompras();
			miPresupC.MdiParent = this;
			miPresupC.Show();
		}

		private void mnuHerrPresupCAnalisis_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			PresupuestoAnalisisCosto miAnalisis = new PresupuestoAnalisisCosto();
			miAnalisis.MdiParent = this;
			miAnalisis.Show();
		}

		private void btConciliacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			NoDeducible miDeducible = new NoDeducible();
			miDeducible.MdiParent = this;
			miDeducible.Show();
			Cursor = Cursors.Default;
		}

		private void mnuHerrPresupCAvanceObra_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			PresupuestoAvanceObra miAvance = new PresupuestoAvanceObra();
			miAvance.MdiParent = this;
			miAvance.Show();
			Cursor = Cursors.Default;
		}

		private void mnuHerrPresupCImportar_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			PresupuestoCImportar miImporta = new PresupuestoCImportar();
			miImporta.MdiParent = this;
			miImporta.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactResJarComision_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			JarContratos miContrato = new JarContratos();
			miContrato.MdiParent = this;
			miContrato.Show();
			Cursor = Cursors.Default;
		}

		private void barEditItem1_EditValueChanged(object sender, System.EventArgs e)
		{
			if (barSkin.EditValue.ToString().Length > 0)
				skinPrincipal.LookAndFeel.SetSkinStyle(barSkin.EditValue.ToString());
		}

		private void mnuHerrSegGenGrupos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			DBUsuarioSeguridad miUsuarioG = new DBUsuarioSeguridad();
			//			miUsuarioG.MdiParent = this;
			//			miUsuarioG.Show();
		}

		private void MnuEdicSegConfigUsuarios_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			DBConfiguracion miAcceso = new DBConfiguracion();
			miAcceso.MdiParent = this;
			miAcceso.Show();
		}

		private void MnuFactContometros_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Contometros miContometro = new Contometros();
			miContometro.MdiParent = this;
			miContometro.Show();
			Cursor = Cursors.Default;
		}

		private void mnuArchImpJuanMarcet_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			JuanMarcet  miJuanMarcet = new JuanMarcet();
			miJuanMarcet.MdiParent = this;
			miJuanMarcet.Show();
		}

		private void mnuArchImpFactRet_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ExcelFacturaRet miFactRet = new ExcelFacturaRet();
			miFactRet.MdiParent = this;
			miFactRet.Show();
		}

		private void mnuHerrAgenda_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Alarma miAlarma = new Alarma();
			miAlarma.MdiParent = this;
			miAlarma.Show();
		}

		private void mnuHerrPresupCReportes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			PresupuestoCReporte miRepP = new PresupuestoCReporte();
			miRepP.MdiParent = this;
			miRepP.Show();
		}

		private void mnuHerrPresupCSolicitudPago_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			PresupuestoOrdenTrabajo miOrdenT = new PresupuestoOrdenTrabajo();
			miOrdenT.MdiParent = this;
			miOrdenT.Show();
		}

		private void mnuAyuJuego_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			PiedraPapel miJuego = new PiedraPapel();
			miJuego.MdiParent = this;
			miJuego.Show();
		}

		private void mnuHerrContratosAmortiza_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			IngresoContratosAmotiza miAmort = new IngresoContratosAmotiza();
			miAmort.MdiParent = this;
			miAmort.Show();
			Cursor = Cursors.Default;		
		}

		private void mnuArchPropEmpresaR_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Empresas miEmpresas = new Empresas();
			miEmpresas.MdiParent = this;
			miEmpresas.Show();
			Cursor = Cursors.Default;		
		}

	
		private void MnuHerrAlarma_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Alarma miAlarma = new Alarma();
			miAlarma.MdiParent = this;
			miAlarma.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFacSpif_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Point_Spif miSpif = new Point_Spif();
			miSpif.MdiParent = this;
			miSpif.Show();
			Cursor = Cursors.Default;
		}

		private void MnnuFactLotesPromociones_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Promociones Promociones = new Promociones();
			Promociones.MdiParent = this;
			Promociones.Show();			
			Cursor = Cursors.Default;
		}

		private void barButtonItem6_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
		
		}

		private void MnuCreditoSolRapida_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
//			Point_IRSolicitud miSolicitudR = new Point_IRSolicitud();
//			miSolicitudR.MdiParent = this;
//			miSolicitudR.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactLoteTransMult_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//			Point_TrnsfMultiple MiTransferenciaM = new Point_TrnsfMultiple();
			//			MiTransferenciaM.MdiParent = this;
			//			MiTransferenciaM.Show();
			frmTransferenciasMultiples Transf = new frmTransferenciasMultiples();
			Transf.MdiParent = this;
			Transf.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactLoteSpif_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			Point_Spif MiSpif = new Point_Spif();
			MiSpif.MdiParent = this;
			MiSpif.Show();
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			lblMensaje.EditValue = "Tiempo: " + ts.TotalSeconds.ToString("0.00");
			Cursor = Cursors.Default;
		}

		private void MnuFacInv_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			Inventarios miInventario = new Inventarios();
			miInventario.MdiParent = this;
			miInventario.Show();
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			lblMensaje.EditValue = "Tiempo: " + ts.TotalSeconds.ToString("0.00");
			Cursor = Cursors.Default;
		}

		private void MnuFactPreFactura_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			try
			{
				//IngresaCompra(43);
				frmPrefactura miPrefactura = new frmPrefactura();
				miPrefactura.MdiParent = this;
				miPrefactura.Show();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "PreFactura");
			}
		}

		private void MnuCreditoCotizador_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;

			Cotizador Cotizador = new Cotizador();
			Cotizador.MdiParent = this;
			Cotizador.Show();

			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			lblMensaje.EditValue = "Tiempo: " + ts.TotalSeconds.ToString("0.00");
			Cursor = Cursors.Default;
		}

		private void MnuFactLotesPromVen_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;

			//PromocionVendedores Promo = new PromocionVendedores();
			frmSpifVendedores SpifVendedores = new frmSpifVendedores();
			SpifVendedores.MdiParent = this;
			SpifVendedores.Show();
			
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			lblMensaje.EditValue = "Tiempo: " + ts.TotalSeconds.ToString("0.00");
			Cursor = Cursors.Default;
		}

		private void MnuCreditoCuota_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Cuota_Point miCuota = new Cuota_Point();
			miCuota.MdiParent = this;
			miCuota.Show();
			Cursor = Cursors.Default;
		}

		private void MnuCreGastoCobran_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			GastosDeCobranza miGasto = new GastosDeCobranza();
			miGasto.MdiParent = this;
			miGasto.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactCombos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmCombos Combos = new frmCombos();
			Combos.MdiParent = this;
			Combos.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactAnticipos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmAnticiposClientes AnticiposClientes = new frmAnticiposClientes();
			AnticiposClientes.MdiParent = this;
			AnticiposClientes.Show();
			Cursor = Cursors.Default;
		}

		private void MnuHerrFormaPago_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			CompraFormaPago miFormaPago = new CompraFormaPago();
			miFormaPago.ShowDialog();
		}

		private void MnuCreCobranza_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCobranzas miCobranza = new frmCobranzas();
			miCobranza.MdiParent = this;
			miCobranza.Show();
		}

		private void MnuFacComConfiguracion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			Point_ConfComision miConf = new Point_ConfComision();
			miConf.MdiParent = this;
			miConf.Show();
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			lblMensaje.EditValue = "Tiempo: " + ts.TotalSeconds.ToString("0.00");
			Cursor = Cursors.Default;
		}

		private void MnuFacArtServicios_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			frmArticulosServicios miArtSer = new frmArticulosServicios();
			miArtSer.MdiParent = this;
			miArtSer.Show();
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			lblMensaje.EditValue = "Tiempo: " + ts.TotalSeconds.ToString("0.00");
			Cursor = Cursors.Default;
		}

		private void MnuFactSerie_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//Seire_Point ERM = new Seire_Point();
			frmEntregaRecepcionMercaderia ERM = new frmEntregaRecepcionMercaderia();
			ERM.MdiParent = this;
			ERM.Show();
			Cursor = Cursors.Default;
		}

		private void mnuFactBodegas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//			Bodegas miBodega = new Bodegas();
			//			miBodega.MdiParent = this;
			//			miBodega.Show();

			frmBodegas Bodega = new frmBodegas();
			Bodega.MdiParent = this;
			Bodega.Show();

			Cursor = Cursors.Default;
		}

		private void MnuFactConsignacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;			
			frmConsignaciones Consignacion = new frmConsignaciones();
			Consignacion.MdiParent = this;
			Consignacion.Show();			
			Cursor = Cursors.Default;
		}

		private void MnuFacPagosAut_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			PagosPendientes miPagoP = new PagosPendientes();
			miPagoP.MdiParent = this;
			miPagoP.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactAutPagos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAutorizacionPagos Autorizacion = new frmAutorizacionPagos();
			Autorizacion.MdiParent = this;
			Autorizacion.Show();

			//			frmAutorizacionDePago Autorizacion = new frmAutorizacionDePago();
			//			Autorizacion.MdiParent = this;
			//			Autorizacion.Show();
		}

		private void MnuContAsigArt_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmArtTipoCuenta miArtCuenta = new frmArtTipoCuenta();
			miArtCuenta.MdiParent=this;
			miArtCuenta.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFacSeriales_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			frmSeriales miSerial = new frmSeriales();
			miSerial.MdiParent = this;
			miSerial.Show();
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			lblMensaje.EditValue = "Tiempo: " + ts.TotalSeconds.ToString("0.00");
			Cursor = Cursors.Default;
		}

		private void MnuContMantTarj_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmMantenimientoTarjetas miMantenimiento = new frmMantenimientoTarjetas();
			miMantenimiento.MdiParent = this;
			miMantenimiento.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactEstTransf_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmEstadoTransferencias miEstado = new frmEstadoTransferencias();
			miEstado.MdiParent = this;
			miEstado.Show();
			Cursor = Cursors.Default;
		}

		private void FacGuiaRemision_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			
		}

		private void mnuFactGuiaRemision_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//frmGuiaDeRemision miGuia = new frmGuiaDeRemision();
			//miGuia.MdiParent = this;
			//miGuia.Show();
			Cursor = Cursors.Default;
		}

		private void mnuFactActPrecio_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmActualizacionDePrecios miActualizacion = new frmActualizacionDePrecios();
			miActualizacion.MdiParent = this;
			miActualizacion.Show();
			Cursor = Cursors.Default;
		}

		private void MnuProdEstado_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmEstadoOrdenProduccion miOrdenEstado = new frmEstadoOrdenProduccion();
			miOrdenEstado.MdiParent = this;
			miOrdenEstado.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFacSerie_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
		
		}

		private void MnuContDepositoVouchers_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			Cursor = Cursors.WaitCursor;
			//			frmCapturaVouchers Captura = new frmCapturaVouchers();
			//			Captura.MdiParent = this;
			//			Captura.Show();

			frmDepositoDeVouchers Depositos = new frmDepositoDeVouchers();
			Depositos.MdiParent = this;
			Depositos.Show();
			Cursor = Cursors.Default;
		}

		private void mnuCreLotesUnibanco_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmLoteClientesCrediPoint miLote = new frmLoteClientesCrediPoint();
			miLote.MdiParent = this;
			miLote.Show();
			Cursor = Cursors.Default;
		}

		private void MnuContCuadreCaja_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//			DateTime dtInicio = DateTime.Now;
			frmCierreDeCaja miCierre = new frmCierreDeCaja();
			miCierre.MdiParent = this;
			miCierre.Show();
			//			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			//			lblMensaje.EditValue = "Tiempo: " + ts.TotalSeconds.ToString("0.00");
			Cursor = Cursors.Default;
		}

		private void MnuContCXP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCuentasPorPagar miCuenta = new frmCuentasPorPagar();
			miCuenta.MdiParent = this;
			miCuenta.Show();
		}

		private void MnuContCXC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCuentasPorCobrar miCuenta = new frmCuentasPorCobrar();
			miCuenta.MdiParent = this;
			miCuenta.Show();
		}

		private void mnuFactVentasConsignacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteVentasConsignacion miVenta = new frmReporteVentasConsignacion();
			miVenta.MdiParent = this;
			miVenta.Show();
		}

		private void MnuEdicCopi_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
		
		}

		private void MenuLatinium_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			//if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void MnuDTSTA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmServiciosTecnicos miST = new frmServiciosTecnicos();
			miST.MdiParent = this;
			miST.Show();
		}

		private void MnuDTSolicitud_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDTIngreso IngresoDT = new frmDTIngreso();
			IngresoDT.MdiParent = this;
			IngresoDT.Show();
		}

		private void mnuArchPropProd_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
		
		}

		private void mnuProduccionPendientes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmDptoTecnico miDptoTecnico = new frmDptoTecnico();
			miDptoTecnico.MdiParent = this;
			miDptoTecnico.Show();
			Cursor = Cursors.Default;
		}

		private void mnuVentasResumenDiario_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//			RepVentas miResumen = new RepVentas();
			//			miResumen.MdiParent = this;
			//			miResumen.Show();

			frmResumenDeVentas Resumen = new frmResumenDeVentas();
			Resumen.MdiParent = this;
			Resumen.Show();
			Cursor = Cursors.Default;
		}

		private void mnuVentasCB_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmCodigosBarras Codigo = new frmCodigosBarras();
			Codigo.MdiParent = this;
			Codigo.Show();
			Cursor = Cursors.Default;
		}

		private void mnuVentasCotizacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmCotizaciones Cotizacion = new frmCotizaciones();
			Cotizacion.MdiParent = this;
			Cotizacion.Show();			
			Cursor = Cursors.Default;
		}

		private void mnuVentaAutDistribuidor_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmAutorizaDistribuidor Autorizacion = new frmAutorizaDistribuidor();
			Autorizacion.MdiParent = this;
			Autorizacion.Show();
			Cursor = Cursors.Default;
		}

		private void mnuVentaLiqCons_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmLiquidaciónConsignacion miLiquidacion = new frmLiquidaciónConsignacion();
			miLiquidacion.MdiParent = this;
			miLiquidacion.Show();
			Cursor = Cursors.Default;
		}

		private void mnuComprasComboZOOM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmGruposDeArticulos miComboZOOM = new frmGruposDeArticulos();
			miComboZOOM.MdiParent = this;
			miComboZOOM.Show();
			Cursor = Cursors.Default;
		}

		private void mnuComprasLiquidaGastos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmGastosPorServicio miGasto = new frmGastosPorServicio();
			miGasto.MdiParent = this;
			miGasto.Show();
			Cursor = Cursors.Default;
		}

		private void mnuMatrizDsctos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmMatrizDescuentosZoom miMatriz = new frmMatrizDescuentosZoom();
			miMatriz.MdiParent = this;
			miMatriz.Show();
			Cursor = Cursors.Default;
		}

		private void mnuComprasAgenda_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmAgenda miAgenda = new frmAgenda();
			miAgenda.MdiParent = this;
			miAgenda.Show();
			Cursor = Cursors.Default;
		}

		private void mnuCompraGestion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmGestionCliente miGestion = new frmGestionCliente();
			miGestion.MdiParent = this;
			miGestion.Show();
			Cursor = Cursors.Default;
		}

		private void mnuVentasComDetallePersonal_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DetallePersonalComisiones miDetalle = new DetallePersonalComisiones();
			miDetalle.MdiParent = this;
			miDetalle.Show();
			Cursor = Cursors.Default;
		}

		private void mnuContResCXC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			Cursor = Cursors.WaitCursor;
			//			frmReporteCuentasPorCobrar miResumen = new frmReporteCuentasPorCobrar();
			//			miResumen.MdiParent = this;
			//			miResumen.Show();

			frmResumenCXC Resumen = new frmResumenCXC();
			Resumen.MdiParent = this;
			Resumen.Show();
			Cursor = Cursors.Default;
		}

		private void mnuContResCXP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmReporteCuentasPorPagar miResumen = new frmReporteCuentasPorPagar();
			miResumen.MdiParent = this;
			miResumen.Show();
			Cursor = Cursors.Default;
		}

		private void mnuContRepAnt_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmReposicionCarteraSolidario Reposicion = new frmReposicionCarteraSolidario();
			Reposicion.MdiParent = this;
			Reposicion.Show();

			//			frmReporteAnticipos miReporteAnticipos = new frmReporteAnticipos();
			//			miReporteAnticipos.MdiParent = this;
			//			miReporteAnticipos.Show();
			Cursor = Cursors.Default;
		}

		private void mnuBodegaKardex_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmMovimientosKardex Kardex = new frmMovimientosKardex();
			Kardex.MdiParent = this;
			Kardex.Show();
			Cursor = Cursors.Default;
		}

		private void mnuRolFiniquitos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmFiniquitosAnticiposViaje miFiniquito = new frmFiniquitosAnticiposViaje();
			miFiniquito.MdiParent = this;
			miFiniquito.Show();
			Cursor = Cursors.Default;
		}

		private void mnuHerramientasNumeracion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmNumeracionLocales Numeracion = new frmNumeracionLocales();
			Numeracion.MdiParent = this;
			Numeracion.Show();
			Cursor = Cursors.Default;
		}

		private void mnuContCCxCCP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmGeneraCruceCxCCrediPoint Cruce = new frmGeneraCruceCxCCrediPoint();
			Cruce.MdiParent = this;
			Cruce.Show();
			Cursor = Cursors.Default;
		}

		private void MnuFactComConsulta_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmConsultaComisiones ConsultaComisiones = new frmConsultaComisiones();
			ConsultaComisiones.MdiParent = this;
			ConsultaComisiones.Show();
		}

	

		private void mnuVentasReporteCompara_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRepVenComparaLatEas Compara = new frmRepVenComparaLatEas();
			Compara.MdiParent = this;
			Compara.Show();
		}

		private void mnuContabilidadAsDesc_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAsientosDescuadrados AsientosDesc = new frmAsientosDescuadrados();
			AsientosDesc.MdiParent = this;
			AsientosDesc.Show();
		}

		private void mnuContAntProv_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAnticiposProveedor Anticipos = new frmAnticiposProveedor();
			Anticipos.MdiParent = this;
			Anticipos.Show();
		}

		private void mnuContRepCC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteCajaChica RCC = new frmReporteCajaChica();
			RCC.MdiParent = this;
			RCC.Show();
		}

		private void mnuContASF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			FrmAnticiposProvSinFactura ASF = new FrmAnticiposProvSinFactura();
			ASF.MdiParent = this;
			ASF.Show();
		}

		private void mnuContMovBanc_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmMovimientosBancarios Movimientos = new frmMovimientosBancarios();
			Movimientos.MdiParent = this;
			Movimientos.Show();
		}

		private void mnuContEstadoCheques_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDetalleCheques Estado = new frmDetalleCheques();
			Estado.MdiParent = this;
			Estado.Show();
		}

		private void mnuComprasRepSpifVend_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteSpifVendedores Reporte = new frmReporteSpifVendedores();
			Reporte.MdiParent = this;
			Reporte.Show();
		}

		private void mnuContPrestamos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPrestamos Prestamos = new frmPrestamos();
			Prestamos.MdiParent = this;
			Prestamos.Show();
		}

		private void mnuContTEB_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmTransferenciasBancarias Transferencia = new frmTransferenciasBancarias();
			Transferencia.MdiParent = this;
			Transferencia.Show();
		}

		private void mnuVentasDA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDevolucionAnticipo DA = new frmDevolucionAnticipo();
			DA.MdiParent = this;
			DA.Show();
		}

		private void mnuContConsumosTC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmGastosTarjetasDeCredito GTC = new frmGastosTarjetasDeCredito();
			GTC.MdiParent = this;
			GTC.Show();
		}

		private void mnuBodegaDCI_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmConteoInventarios DCI = new frmConteoInventarios();
			DCI.MdiParent = this;
			DCI.Show();
		}

		private void mnuContLCTC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmConsumosTarjetasDeCredito LCTC = new frmConsumosTarjetasDeCredito();
			LCTC.MdiParent = this;
			LCTC.Show();
		}

		private void mnuProduccionTPT_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmTransferenciaProduccion TransferenciaProduccion = new frmTransferenciaProduccion();
			TransferenciaProduccion.MdiParent = this;
			TransferenciaProduccion.Show();
		}

		private void mnuContTCC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmTarjetasCorporativas TCC = new frmTarjetasCorporativas();
			TCC.MdiParent = this;
			TCC.Show();
		}

		private void barButtonItem14_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			CRMArticulo art = new CRMArticulo();
			art.MdiParent = this;
			art.Show();

			CRMClientes cli = new CRMClientes();
			cli.MdiParent = this;
			cli.Show();
		}

		private void mnuDTIngresoMultiple_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmIngresoDepartamentoTecnico IDT = new frmIngresoDepartamentoTecnico();
			IDT.MdiParent = this;
			IDT.Show();
		}

		private void mnuDTRetiro_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDTIngresoPorRetiro IDTR = new frmDTIngresoPorRetiro();
			IDTR.MdiParent = this;
			IDTR.Show();
		}

		private void mnuBodegaCI_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//frmRepComparaInvEasyLat Compara = new frmRepComparaInvEasyLat();
			frmConteoDeInventario Conteo = new frmConteoDeInventario();
			Conteo.MdiParent = this;
			Conteo.Show();
		}

		private void mnuVentasImpresionDePrecios_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmImpresionDePrecios ImprimePrecios = new frmImpresionDePrecios();
			ImprimePrecios.MdiParent = this;
			ImprimePrecios.Show();
		}

		private void mnuHerrSegCD_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmClaveDescuento CD = new frmClaveDescuento();
			CD.MdiParent = this;
			CD.Show();
		}

		private void mnuRepCmpTransf_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDeTransferencias RDT = new frmReporteDeTransferencias();
			RDT.MdiParent = this;
			RDT.Show();
		}

		private void barButtonItem12_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAjustaCXCXP Ajuste = new frmAjustaCXCXP();
			Ajuste.MdiParent = this;
			Ajuste.Show();

			frmAjustaCXC Ajuste1 = new frmAjustaCXC();
			Ajuste1.MdiParent = this;
			Ajuste1.Show();
		}

		private void mnuBodegaCierreInventario_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmConteoJustificaDiferencias Justifica = new frmConteoJustificaDiferencias();
			Justifica.MdiParent = this;
			Justifica.Show();
		}

		private void barButtonItem17_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmFCEasyLat FactEM = new frmFCEasyLat();
			FactEM.MdiParent = this;
			FactEM.Show();
		}

		private void barButtonItem18_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmProduccionPortatiles ProdPort = new frmProduccionPortatiles();
			ProdPort.MdiParent = this;
			ProdPort.Show();
		}

		private void mnuGPPD_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmSolicitaTransferenciasPartesDañadasProduccion Transferencia = new frmSolicitaTransferenciasPartesDañadasProduccion();
			Transferencia.MdiParent = this;
			Transferencia.Show();
		}

		private void mnuCPReportesCliPoint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteCrediPoint ReporteCrediPoint = new frmReporteCrediPoint();
			ReporteCrediPoint.MdiParent = this;
			ReporteCrediPoint.Show();
		}

		private void mnuCPInformeCobrador_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmInformeCobrador InformesCobrador = new frmInformeCobrador();
			InformesCobrador.MdiParent = this;
			InformesCobrador.Show();
		}

		private void mnuCompraReportesCPA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDeComprasPorProducto CompraPorProducto = new frmReporteDeComprasPorProducto();
			CompraPorProducto.MdiParent = this;
			CompraPorProducto.Show();
		}

		private void mnuVentasReportesPorMarca_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteVentasPorMarca Reportes = new frmReporteVentasPorMarca();
			Reportes.MdiParent = this;
			Reportes.Show();
		}

		private void mnuHerramientasCRIVA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmConfiguraRetencionesIVA Config = new frmConfiguraRetencionesIVA();
			Config.MdiParent = this;
			Config.Show();
		}

		private void mnuComprasPromociones_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDescripcionPromociones Descripcion = new frmDescripcionPromociones();
			Descripcion.MdiParent = this;
			Descripcion.Show();
		}

		private void mnuVentaPromocionesActivas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmVistaPromociones Vista = new frmVistaPromociones();
			Vista.MdiParent = this;
			Vista.Show();
		}

		private void mnuBodegaIF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmInventarioPorFecha IF = new frmInventarioPorFecha();
			IF.MdiParent = this;
			IF.Show();
		}

		private void mnuVentasRVML_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteVentasPorLocalMarca RVML = new frmReporteVentasPorLocalMarca();
			RVML.MdiParent = this;
			RVML.Show();
		}

		private void mnuCrediPointDFG_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDebitosFondoDeGarantia DFG = new frmDebitosFondoDeGarantia();
			DFG.MdiParent = this;
			DFG.Show();
		}

		private void mnuCrediPointMEF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmModificaEntFinanciera MEF = new frmModificaEntFinanciera();
			MEF.MdiParent = this;
			MEF.Show();
		}

		private void mnuCPCP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmListaCompromisosDePago Lista = new frmListaCompromisosDePago();
			Lista.MdiParent = this;
			Lista.Show();
		}

		private void mnuCPIO_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmInformeOperadora Informe = new frmInformeOperadora();
			Informe.MdiParent = this;
			Informe.Show();
		}

		private void mnuBodegaSectorizacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDistanciasBodegas Sectorizacion = new frmDistanciasBodegas();
			Sectorizacion.MdiParent = this;
			Sectorizacion.Show();
		}

		private void mnuComprasTEL_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmTransferenciasExcesoMercaderia TEL = new frmTransferenciasExcesoMercaderia();
			TEL.MdiParent = this;
			TEL.Show();
		}

		private void mnuVentasIPP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmImpresionDePreciosBaratazo ImprimePrecios = new frmImpresionDePreciosBaratazo();
			ImprimePrecios.MdiParent = this;
			ImprimePrecios.Show();
		}

		private void mnuContBancosCBA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmConciliacionAutomatica ConciliacionAutomatica = new frmConciliacionAutomatica();
			ConciliacionAutomatica.MdiParent = this;
			ConciliacionAutomatica.Show();
		}

		private void mnuBodegaInvVal_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmInventarioCostoPorBodega Costo = new frmInventarioCostoPorBodega();
			Costo.MdiParent = this;
			Costo.Show();
		}

		private void mnuCPLCS_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCreditosSolidario CS = new frmCreditosSolidario();
			CS.MdiParent = this;
			CS.Show();
		}

		private void mnuCPRepEquifax_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteEquifax Equifax = new frmReporteEquifax();
			Equifax.MdiParent = this;
			Equifax.Show();
		}

		private void mnuCPMFCP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmActualizaFechaTablaDeAmortizacion AFTA = new frmActualizaFechaTablaDeAmortizacion();
			AFTA.MdiParent = this;
			AFTA.Show();
		}

		private void mnuVentasRepMU_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteMargenUtilidadVentas MUV = new frmReporteMargenUtilidadVentas();
			MUV.MdiParent = this;
			MUV.Show();
		}

		private void mnuContRepGastos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDeGastos CRG = new frmReporteDeGastos();
			CRG.MdiParent = this;
			CRG.Show();
		}

		private void mnuCPRepCP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			
		}

		private void barButtonItem20_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmLoteBancoDelPichincha LBDP = new frmLoteBancoDelPichincha();
			LBDP.MdiParent = this;
			LBDP.Show();
		}

		private void mnuContRepFacElec_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRepFElectronica RFE = new frmRepFElectronica();
			RFE.MdiParent = this;
			RFE.Show();
		}

		private void barButtonItem29_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmHabitaciones LoteHab = new frmHabitaciones();
			LoteHab.MdiParent = this;
			LoteHab.Show();
		}

		private void barButtonItem30_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmTematicas TemaHab = new frmTematicas();
			TemaHab.MdiParent = this;
			TemaHab.Show();		
		}

		private void barButtonItem31_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			tEstadia tEs = new tEstadia();
			tEs.MdiParent = this;
			tEs.Show();		
		}

		private void barButtonItem32_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmTiemposHabitaciones tHabi = new frmTiemposHabitaciones();
			tHabi.MdiParent = this;
			tHabi.Show();		
		}

		private void barButtonItem33_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCuadriculaHabitaciones CuadriHab = new frmCuadriculaHabitaciones();
			CuadriHab.MdiParent = this;
			CuadriHab.Show();		
		}

		private void barButtonItem34_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCobroPrendas cPren = new frmCobroPrendas();
			cPren.MdiParent = this;
			cPren.Show();		
		}

		private void barButtonItem35_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmKardexLotte fKarLo = new frmKardexLotte();
			fKarLo.MdiParent = this;
			fKarLo.Show();		
		}

		private void barButtonItem36_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			frmPromocionesLote PL = new frmPromocionesLote();
			//			PL.MdiParent = this;
			//			PL.Show();	
			frmPromociones PromocionesPeony = new frmPromociones();
			PromocionesPeony.MdiParent = this;
			PromocionesPeony.Show();
		}

		private void mnuCpLbp_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmLiquidacionBancoPichincha flbco = new frmLiquidacionBancoPichincha();
			flbco.MdiParent = this;
			flbco.Show();		
		}

		private void mnuVentasReportesAIM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteAnalisisInformacionMensual AIM = new frmReporteAnalisisInformacionMensual();
			AIM.MdiParent = this;
			AIM.Show();		
		}

		private void mnuVentasReportesRF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRevisionDeFacturas RF = new frmRevisionDeFacturas();
			RF.MdiParent = this;
			RF.Show();
		}

		private void mnuReservas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReservasHabitaciones RH = new frmReservasHabitaciones();
			RH.MdiParent = this;
			RH.Show();
		}

		private void barButtonItem22_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRepoReservas RS = new frmRepoReservas();
			RS.MdiParent = this;
			RS.Show();
		}

		private void barButtonItem24_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmMovimientosInvLotte MIL = new frmMovimientosInvLotte();
			MIL.MdiParent = this;
			MIL.Show();
		}

		private void barButtonItem25_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRepoCajaLotte RCL = new frmRepoCajaLotte();
			RCL.MdiParent = this;
			RCL.Show();
		}

		private void barButtonItem26_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			frmFacturacion Facturacion = new frmFacturacion();
			//			Facturacion.MdiParent = this;
			//			Facturacion.Show();

			frmFacturacionHabitaciones FH = new frmFacturacionHabitaciones();
			FH.MdiParent = this;
			FH.Show();
		}

		private void mnuVentasVentasCP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmVentaCrediPoint VCP = new frmVentaCrediPoint();
			VCP.MdiParent = this;
			VCP.Show();
		}

		private void barButtonItem27_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmMinimosArt MI = new frmMinimosArt();
			MI.MdiParent = this;
			MI.Show();
		}

		private void mnuHerrALF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAnulaFacturas ALF = new frmAnulaFacturas();
			ALF.MdiParent = this;
			ALF.Show();
		}

		private void mnuBodegaIngSeriales_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmIngresoManualDeSeriales IMS = new frmIngresoManualDeSeriales();
			IMS.MdiParent = this;
			IMS.Show();
		}

		private void barButtonItem28_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAjustesInventarios fAjsInv = new frmAjustesInventarios();
			fAjsInv.MdiParent = this;
			fAjsInv.Show();
		}

		private void barButtonItem29_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmTransferLotte fTrLot = new frmTransferLotte();
			fTrLot.MdiParent = this;
			fTrLot.Show();
		}

		private void MnuFactComCalificacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
		
		}

		private void MnuFactDetalle_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
		
		}

		private void barButtonItem30_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmTurnosLotte fTurns = new frmTurnosLotte();
			fTurns.MdiParent = this;
			fTurns.Show();
		}

		private void mnuBodegaGRC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmGuiaDeRemisionCarpa GRC = new frmGuiaDeRemisionCarpa();
			GRC.MdiParent = this;
			GRC.Show();
		}

		private void barButtonItem31_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRepDiarioVentasLotte fRDVLot = new frmRepDiarioVentasLotte();
			fRDVLot.MdiParent = this;
			fRDVLot.Show();
		}

		private void barButtonItem32_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRepResumenVentasLotte fRDVLot = new frmRepResumenVentasLotte();
			fRDVLot.MdiParent = this;
			fRDVLot.Show();
		}

		private void mnuCPRRCS_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteAnticiposSolidario RAS = new frmReporteAnticiposSolidario();
			RAS.MdiParent = this;
			RAS.Show();
		}

		private void mnuComprasGDA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmGruposArticulos GDA = new frmGruposArticulos();
			GDA.MdiParent = this;
			GDA.Show();
		}

		private void mnuLoteNotasDeEntrega_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//frmNotasEntrega NDE = new frmNotasEntrega();
			frmNotaDeEntrega NDE = new frmNotaDeEntrega();
			NDE.MdiParent = this;
			NDE.Show();
		}

		private void mnuCPRD_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_RecepcionDocumentos CRD = new frmCre_RecepcionDocumentos();
			CRD.MdiParent = this;
			CRD.Show();
		}

		private void mnuCPRepRDD_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ReporteRecepcionDocumentos CRRD = new frmCre_ReporteRecepcionDocumentos();
			CRRD.MdiParent = this;
			CRRD.Show();
		}

		private void mnuCPVEP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_VerificacionEntregaProductos CVEP = new frmCre_VerificacionEntregaProductos();
			CVEP.MdiParent = this;
			CVEP.Show();
		}

		private void mnuVentasARP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmActasDeRecibido ARP = new frmActasDeRecibido();
			ARP.MdiParent = this;
			ARP.Show();
		}

		private void mnuCPRepECP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ReporteEstadoCrediPoint RECP = new frmCre_ReporteEstadoCrediPoint();
			RECP.MdiParent = this;
			RECP.Show();
		}

		private void mnuCPEDC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_EnvioCedulasVendedores CECV = new frmCre_EnvioCedulasVendedores();
			CECV.MdiParent = this;
			CECV.Show();
		}

		private void mnuComprasRepFP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmFlujoProveedores FP = new frmFlujoProveedores();
			FP.MdiParent = this;
			FP.Show();
		}

		private void barButtonItem34_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPresupuestoPorLocal Presupuesto = new frmPresupuestoPorLocal();
			Presupuesto.MdiParent = this;
			Presupuesto.Show();
		}

		private void barButtonItem35_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			frmAsignaciondeVendedores Asignacion = new frmAsignaciondeVendedores();
			//			Asignacion.MdiParent = this;
			//			Asignacion.Show();
			//			frmMovimientoVendedores Movimiento = new frmMovimientoVendedores();
			//			Movimiento.MdiParent = this;
			//			Movimiento.Show();
		}

		private void barButtonItem36_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmXDescuento fxDes = new frmXDescuento();
			fxDes.MdiParent = this;
			fxDes.Show();
		}

		private void mnuCPHRC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ReporteHistorialRevisionDeCedulas Historial = new frmCre_ReporteHistorialRevisionDeCedulas();
			Historial.MdiParent = this;
			Historial.Show();
		}

		private void mnuVentaRMU_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteMargenUtilidadPorLocal RMUL = new frmReporteMargenUtilidadPorLocal();
			RMUL.MdiParent = this;
			RMUL.Show();

			//			frmResumenMargenUtilidad RMU = new frmResumenMargenUtilidad();
			//			RMU.MdiParent = this;
			//			RMU.Show();
		}

		private void mnuCPRRC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ResumenRevisionDeCedulas RRDC = new frmCre_ResumenRevisionDeCedulas();
			RRDC.MdiParent = this;
			RRDC.Show();
		}

		private void barButtonItem37_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ReporteRetirosDeMercaderia RRDM = new frmCre_ReporteRetirosDeMercaderia();
			RRDM.MdiParent = this;
			RRDM.Show();
		}

		private void mnuDTIRM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmInformeDeRetiroDeMercaderia IRDM = new frmInformeDeRetiroDeMercaderia();
			IRDM.MdiParent = this;
			IRDM.Show();
		}

		private void mnuDTIRMP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmInformeRetiroMercaderiaProductos IRDMP = new frmInformeRetiroMercaderiaProductos();
			IRDMP.MdiParent = this;
			IRDMP.Show();
		}

		private void mnuRolesCom_AV_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			frmCom_AsignacionVendedoresSemanal Asignacion = new frmCom_AsignacionVendedoresSemanal();
			//			Asignacion.MdiParent = this;
			//			Asignacion.Show();

			frmCom_AsignacionVendedores Asignacion = new frmCom_AsignacionVendedores();
			Asignacion.MdiParent = this;
			Asignacion.Show();
		}

		private void mnuRolesCom_PA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_PresupuestoEmpresa CPE = new frmCom_PresupuestoEmpresa();
			CPE.MdiParent = this;
			CPE.Show();
		}

		private void mnuRolesCom_PML_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_PresupuestoMensualPorLocal CPML = new frmCom_PresupuestoMensualPorLocal();
			CPML.MdiParent = this;
			CPML.Show();
		}

		private void mnuRolesCom_PL_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_PersonalPorLocal CPL = new frmCom_PersonalPorLocal();
			CPL.MdiParent = this;
			CPL.Show();
		}

		private void mnuVentasPC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPlanificacionHorariosDeCarpa PHC = new frmPlanificacionHorariosDeCarpa();
			PHC.MdiParent = this;
			PHC.Show();
			//frmHorariosDeCarpa HDC = new frmHorariosDeCarpa();
			//			HDC.MdiParent = this;
			//			HDC.Show();
		}

		private void barButtonItem35_ItemClick_2(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCajaGastos CG = new frmCajaGastos();
			CG.MdiParent = this;
			CG.Show();
		}

		private void mnuRolPA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			RolPersonal miPersonal = new RolPersonal();
			miPersonal.MdiParent = this;
			miPersonal.Show();		
		}

		private void mnuCPCobranzasRCP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ResumenCarteraPichincha RCP = new frmCre_ResumenCarteraPichincha();
			RCP.MdiParent = this;
			RCP.Show();
		}

		private void mnuRolComisionesPC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			frmCom_PorcentajesComisionesSemanales PCS = new frmCom_PorcentajesComisionesSemanales();
			//			PCS.MdiParent = this;
			//			PCS.Show();
			//			frmCom_PorcentajesComisiones CP = new frmCom_PorcentajesComisiones();
			//			CP.MdiParent = this;
			//			CP.Show();
			frmReporteComisiones RC = new frmReporteComisiones();
			RC.MdiParent = this;
			RC.Show();
		}

		private void barButtonItem40_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRepVentasLotte rvl = new frmRepVentasLotte();
			rvl.MdiParent = this;
			rvl.Show();
		}

		private void mnuVentasResumenComisiones_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			if (MenuLatinium.iNivel == 4)
			//			{
			//			frmCom_ResumenPorVendedor CRV = new frmCom_ResumenPorVendedor();
			//			CRV.MdiParent = this;
			//			CRV.Show();
			//			}

			//				if (MenuLatinium.iNivel == 1)
			//			{		
			frmCom_ResumenMensual RCS = new frmCom_ResumenMensual();
			RCS.MdiParent = this;
			RCS.Show();
			//									}
			
			//			if (MenuLatinium.iNivel == 1)
			//			{
			//				frmReporte RT = new frmReporte();
			//				RT.MdiParent = this;
			//				RT.Show();
			//				//			else
			//				//			{
			//				//				frmCom_ResumenDeComisionesSemanal RSM = new frmCom_ResumenDeComisionesSemanal();
			//				//				RSM.MdiParent = this;
			//				//				RSM.Show();
			//			}
		}

		private void mnuCPReportesRCV_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ReporteCarteraVendida CRCV = new frmCre_ReporteCarteraVendida();
			CRCV.MdiParent = this;
			CRCV.Show();
		}

		private void mnuRolesPagoComisiones_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_PagoPersonal CPP = new frmCom_PagoPersonal (bComisionQuincena);
			CPP.MdiParent = this;
			CPP.Show();
		}

		private void mnuRolesComAJV_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_AsignacionesJefesDeVentas CAJV = new frmCom_AsignacionesJefesDeVentas();
			//frmCom_AsignacionSemanalJefes CAJV = new frmCom_AsignacionSemanalJefes();
			CAJV.MdiParent = this;
			CAJV.Show();
		}

		private void mnuCPDFGP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmFondoGarantiaPichincha DFGP = new frmFondoGarantiaPichincha();
			DFGP.MdiParent = this;
			DFGP.Show();
		}

		private void mnuComprasAPC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReportesArticulosPorComprar RAPC = new frmReportesArticulosPorComprar();
			RAPC.MdiParent = this;
			RAPC.Show();
		}

		private void mnuRolesReportesPVF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmNominaReporteSalidas NRS = new frmNominaReporteSalidas();
			NRS.MdiParent = this;
			NRS.Show();
		}

		private void barButtonItem42_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmResmenVentasLotte RVL = new frmResmenVentasLotte();
			RVL.MdiParent = this;
			RVL.Show();
		}

		private void mnuCom_PAGE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_PresupuestoEmpresaGE PEGE = new frmCom_PresupuestoEmpresaGE();
			PEGE.MdiParent = this;
			PEGE.Show();
		}

		private void mnuCom_PMLGE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_PresupuestoMensualPorLocalGE PMLGE = new frmCom_PresupuestoMensualPorLocalGE();
			PMLGE.MdiParent = this;
			PMLGE.Show();
		}

		private void mnuCom_PCGE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_PorcentajesComisionesGE PCGE = new frmCom_PorcentajesComisionesGE();
			PCGE.MdiParent = this;
			PCGE.Show();
		}

		private void mnuVentasRepGE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteGarantiaExtendida RGE = new frmReporteGarantiaExtendida();
			RGE.MdiParent = this;
			RGE.Show();
		}

		private void mnuCom_ReportesRM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_ResumenComisiones RC = new frmCom_ResumenComisiones();
			RC.MdiParent = this;
			RC.Show();
		}

		private void mnuCobranzasRGC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ReporteGestionCobranza RGC = new frmCre_ReporteGestionCobranza();
			RGC.MdiParent = this;
			RGC.Show();
		}

	
		private void mnuRolesReportesVXM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_ReporteVendedores RV = new frmCom_ReporteVendedores();
			RV.MdiParent = this;
			RV.Show();
		}

		private void mnuVentasMarcador_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{			
			frmMarcador Marcador = new frmMarcador();
			//frmMarcadorGrupo Marcador = new frmMarcadorGrupo();
			Marcador.MdiParent = this;
			Marcador.Show();
		}

		private void mnuCPLotesPPagos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_TramaDePagosBancoPichincha Pagos = new frmCre_TramaDePagosBancoPichincha();
			Pagos.MdiParent = this;
			Pagos.Show();
		}

		private void mnuVentasReportesComparativos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//frmReportesDeVentas RDV = new frmReportesDeVentas();
			frmComparativosDeVentas CDV = new frmComparativosDeVentas();
			CDV.MdiParent = this;
			CDV.Show();
		}

		private void mnuCPReportesGDCM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmGestionDeCobranzaMensual GCM = new frmGestionDeCobranzaMensual();
			GCM.MdiParent = this;
			GCM.Show();
		}

		private void mnuBodegaIPR_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDTCreaProductoDeRemate CPR = new frmDTCreaProductoDeRemate();
			CPR.MdiParent = this;
			CPR.Show();
		}

		private void barButtonItem50_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmComparaCobranzaPichincha CCP = new frmComparaCobranzaPichincha();
			CCP.MdiParent = this;
			CCP.Show();
		}

		private void barButtonItem51_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAsignacionTecnicos fatec =  new frmAsignacionTecnicos();
			fatec.MdiParent = this;
			fatec.Show();		
		}

		private void barButtonItem52_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{			
			frmEdoProductosDT fedopd =  new frmEdoProductosDT();
			fedopd.MdiParent = this;
			fedopd.Show();	
		}

		private void barButtonItem53_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReparacionProductosDT frepd =  new frmReparacionProductosDT();
			frepd.MdiParent = this;
			frepd.Show();	
		}

		private void barButtonItem54_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmEntregaProductosDT fentpd =  new frmEntregaProductosDT();
			fentpd.MdiParent = this;
			fentpd.Show();	
		}

		private void mnuContRepCajasChicas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteCajasChicas RCC = new frmReporteCajasChicas();
			RCC.MdiParent = this;
			RCC.Show();
		}

		private void mnuCPReportesAbonos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCxCReporteAnticipos RA = new frmCxCReporteAnticipos();
			RA.MdiParent = this;
			RA.Show();
		}

		private void mnuCPReportesNDC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCXCReporteNotasDeCredito RNDC = new frmCXCReporteNotasDeCredito();
			RNDC.MdiParent = this;
			RNDC.Show();
		}

		private void mnuVentasReportesVentasPorMes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDeVentasPorMes RVPM = new frmReporteDeVentasPorMes();
			RVPM.MdiParent = this;
			RVPM.Show();
		}

		private void mnuContReportesGA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDeGastosAnual RGA = new frmReporteDeGastosAnual();
			RGA.MdiParent = this;
			RGA.Show();
		}

		private void mnuContBancosCP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteCheques RCP = new frmReporteCheques();
			RCP.MdiParent = this;
			RCP.Show();
		}

		private void mnuVentasReportesTC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteTarjetasVentas RTV = new frmReporteTarjetasVentas();
			RTV.MdiParent = this;
			RTV.Show();
		}

		private void mnuRolesReportesUniformes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmNominaReportes NRU = new frmNominaReportes();
			NRU.MdiParent = this;
			NRU.Show();
		}

		private void barButtonItem57_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCarteraSolidario CS = new frmCarteraSolidario();
			CS.MdiParent = this;
			CS.Show();
		}

		private void mnuContBancosDPL_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDepositosPorLiquidar DPL = new frmDepositosPorLiquidar();
			DPL.MdiParent = this;
			DPL.Show();
		}

		private void mnuContBancosCDPL_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCruzaDepositosPorLiquidar CDPL = new frmCruzaDepositosPorLiquidar();
			CDPL.MdiParent = this;
			CDPL.Show();
		}

		private void barButtonItem59_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmXAutorizarDT xautodt = new frmXAutorizarDT();
			xautodt.MdiParent = this;
			xautodt.Show();
		}

		private void barButtonItem60_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDTNc dtnc = new frmDTNc();
			dtnc.MdiParent = this;
			dtnc.Show();		
		}

		private void barButtonItem62_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmEstadoProductosDT dtestdo = new frmEstadoProductosDT();
			dtestdo.MdiParent = this;
			dtestdo.Show();			
		}

		private void barButtonItem63_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRepTecnicos reptec = new frmRepTecnicos();
			reptec.MdiParent = this;
			reptec.Show();				
		}

		private void barButtonItem64_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDTGarantia dtgarn = new frmDTGarantia();
			dtgarn.MdiParent = this;
			dtgarn.Show();
		}

		private void barButtonItem65_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDTSeguimiento dts = new frmDTSeguimiento();
			dts.MdiParent = this;
			dts.Show();
		}

		private void barButtonItem66_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			
			frmVentaRemates Ventas = new frmVentaRemates();
			Ventas.MdiParent = this;
			Ventas.Show();
			
		}

		private void mnuVentasRepRVF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteVendedoresPorFecha RVPF = new frmReporteVendedoresPorFecha();
			RVPF.MdiParent = this;
			RVPF.Show();
		}

		private void mnuVentasPlanifVentas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_PlanificacionMensualLocales PML = new frmCom_PlanificacionMensualLocales();
			PML.MdiParent = this;
			PML.Show();
		}

		private void barButtonItem67_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRepVentasArtLote fRVM = new frmRepVentasArtLote();
			fRVM.MdiParent = this;
			fRVM.Show();
		}

		private void mnuDTProductosDeFeria_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmProductosDeFeria PDF = new frmProductosDeFeria();
			PDF.MdiParent = this;
			PDF.Show();
		}

		private void barButtonItem69_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRepVentasMesAnio RVMA = new frmRepVentasMesAnio();
			RVMA.MdiParent = this;
			RVMA.Show();
		}

		private void mnuVentasPlanifJefeVentas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPlanificacionVentas PV = new frmPlanificacionVentas();
			PV.MdiParent = this;
			PV.Show();
		}

		private void barButtonItem68_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Form5 f = new Form5();
			f.MdiParent = this;
			f.Show();
		}

		private void mnuRolesReportesNominaGeneral_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteNominaGeneral RNG = new frmReporteNominaGeneral();
			RNG.MdiParent = this;
			RNG.Show();
		}

		private void mnuRolesCom_EPV_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_EvaluacionPersonalVentas EPV = new frmCom_EvaluacionPersonalVentas();
			EPV.MdiParent = this;
			EPV.Show();
		}

		private void mnuRolPersonalPrueba_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmNominaAvisoPersonalPrueba NAPP = new frmNominaAvisoPersonalPrueba();
			NAPP.MdiParent = this;
			NAPP.Show();
		}

		private void mnuRolComPIA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPresupuestoInicioAnio PIA = new frmPresupuestoInicioAnio();
			PIA.MdiParent = this;
			PIA.Show();
		}

		private void mnuVentasReportesDTC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteTarjetasDeCredito RTC = new frmReporteTarjetasDeCredito();
			RTC.MdiParent = this;
			RTC.Show();
		}

		private void mnuVentasReportesEncuestas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReportesFacturacionEncuestas REF = new frmReportesFacturacionEncuestas();
			REF.MdiParent = this;
			REF.Show();
		}

		private void mnuLoteRepProcSIM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmSubeVentas SV = new frmSubeVentas();
			SV.MdiParent = this;
			SV.Show();
		}

		private void mnuCompraCSA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmArticulosCaracteristicas AC = new frmArticulosCaracteristicas();
			AC.MdiParent = this;
			AC.Show();
		}

		private void barButtonItem75_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmComparaDatos CD = new frmComparaDatos();
			CD.MdiParent = this;
			CD.Show();
		}

		private void mnuVentasGCCC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ListaClientesCreditoRevisados CLCCR = new frmCre_ListaClientesCreditoRevisados();
			CLCCR.MdiParent = this;
			CLCCR.Show();
		}

		private void barButtonItem76_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ReporteRetiros RR = new frmCre_ReporteRetiros();
			RR.MdiParent = this;
			RR.Show();
		}

		private void mnuBodegaRM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCalculoRequerimientoInventario CRI = new frmCalculoRequerimientoInventario();
			CRI.MdiParent = this;
			CRI.Show();
		}

		private void mnuCPReportesBR_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_BalanceRecuperacion BR = new frmCre_BalanceRecuperacion();
			BR.MdiParent = this;
			BR.Show();
		}

		private void mnuHerramientasPTP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPlanificacionActividadesPersonal PAP = new frmPlanificacionActividadesPersonal();
			PAP.MdiParent = this;
			PAP.Show();
		}

		private void mnuHerramientasAP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAPAsignacion APA = new frmAPAsignacion();
			APA.MdiParent = this;
			APA.Show();
		}

		private void mnuBodegaRepKS_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmKardexSeriales KS = new frmKardexSeriales();
			KS.MdiParent = this;
			KS.Show();
		}

		private void barButtonItem72_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ListaClientesEnvioMensajesTexto LCEMT = new frmCre_ListaClientesEnvioMensajesTexto();
			LCEMT.MdiParent = this;
			LCEMT.Show();
		}

		private void barButtonItem73_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmFacturasDT frmFDT = new frmFacturasDT();
			frmFDT.MdiParent = this;
			frmFDT.Show();		
		}

		private void barButtonItem77_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmSegimientoProductos frmSPRo = new frmSegimientoProductos();
			frmSPRo.MdiParent = this;
			frmSPRo.Show();		
		}

		private void barButtonItem78_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDTCambioEstados frmCamEs = new frmDTCambioEstados();
			frmCamEs.MdiParent = this;
			frmCamEs.Show();			
		}

		private void mnuRolCom_PorcCarpa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_PorcentajesCarpa CPC = new frmCom_PorcentajesCarpa();
			CPC.MdiParent = this;
			CPC.Show();
		}

		private void mnuReporteVentasMensualesPorLocal_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_ReporteVentasLocalesPorMes RVLM = new frmCom_ReporteVentasLocalesPorMes();
			RVLM.MdiParent = this;
			RVLM.Show();
		}

		private void mnuVentasIPR_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPreciadorRemates PR = new frmPreciadorRemates();
			PR.MdiParent = this;
			PR.Show();
		}

		private void mnuBodegaCDI_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmConteoDiarioInventario CDI = new frmConteoDiarioInventario();
			CDI.MdiParent = this;
			CDI.Show();
		}

		private void frmDTTransfLocales_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDTTransferencia TA = new frmDTTransferencia();
			TA.MdiParent = this;
			TA.Show();
		}

		private void mnuDTRA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRecepcionArticulosDT RA = new frmRecepcionArticulosDT();
			RA.MdiParent = this;
			RA.Show();
		}

		private void mnuDTReporteEstadoIngresos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDTReporteEstadoIngresos REI = new frmDTReporteEstadoIngresos();
			REI.MdiParent = this;
			REI.Show();
		}

		private void mnuCPCobranzasLPCA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ListaAbonosClientesAtrasados LACA = new frmCre_ListaAbonosClientesAtrasados();
			//LACA.MdiParent = this;
			LACA.Show();
		}

		private void mnuVentasReportesVA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteVentasPorArticulos RVPA = new frmReporteVentasPorArticulos();
			RVPA.MdiParent = this;
			RVPA.Show();
		}

		private void mnuDTAprobacionRemates_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDTListaRemates DTLR = new frmDTListaRemates();
			DTLR.MdiParent = this;
			DTLR.Show();
		}

		private void mnuDTASSF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDTAprobacionSalidaSinFacturar ASSF = new frmDTAprobacionSalidaSinFacturar();
			ASSF.MdiParent = this;
			ASSF.Show();
		}

		private void barButtonItem53_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			//			frmPlanificacionPagoProveedor PPP = new frmPlanificacionPagoProveedor();
			//			PPP.MdiParent = this;
			//			PPP.Show();

			frmProvisionesFacturas PF = new frmProvisionesFacturas();
			PF.MdiParent = this;
			PF.Show();
		}
		
		private void barButtonItem60_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCalculoIRP CIRP = new frmCalculoIRP();
			CIRP.MdiParent = this;
			CIRP.Show();
		}

		private void barButtonItem62_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmTiempoContratacion TC = new frmTiempoContratacion();
			TC.MdiParent = this;
			TC.Show();
		}

		private void mnuBodegaReportesECDI_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteConteoDiarioInventario RCDI = new frmReporteConteoDiarioInventario();
			RCDI.MdiParent = this;
			RCDI.Show();
		}

		private void mnuSTA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteServicoTecnico STA = new frmReporteServicoTecnico();
			STA.MdiParent = this;
			STA.Show();
		}

		private void mnuVentasReportesFFP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDeVentasFacturasPorFormasDePago RFFP = new frmReporteDeVentasFacturasPorFormasDePago();
			RFFP.MdiParent = this;
			RFFP.Show();
		}

		private void mnuVentasEOC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmEntregaDeObsequiosClientesCumpleaneros EOCC = new frmEntregaDeObsequiosClientesCumpleaneros();
			EOCC.MdiParent = this;
			EOCC.Show();
		}

		private void mnuLottePlanificacion_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPlanificacionDiariaPersonal PDP = new frmPlanificacionDiariaPersonal();
			PDP.MdiParent = this;
			PDP.Show();		
		}

		private void mnuVentasReportesVPD_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmResumenVentasLocalesPorDia RVPD = new frmResumenVentasLocalesPorDia();
			RVPD.MdiParent = this;
			RVPD.Show();		
		}

		private void mnuVentasGCR_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_GestionesReferidos GCR = new frmCre_GestionesReferidos();
			GCR.MdiParent = this;
			GCR.Show();
		}

		private void mnuLotteReportesIT_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteIngresosPorTurnos RIPT = new frmReporteIngresosPorTurnos();
			RIPT.MdiParent = this;
			RIPT.Show();
		}

		private void barButtonItem73_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRegistraMetasLocalTurno RMLT = new frmRegistraMetasLocalTurno();
			RMLT.MdiParent = this;
			RMLT.Show();
		}

		private void mnuLotteCompetencia_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmResultadoCompetenciaDiaria RCD = new frmResultadoCompetenciaDiaria();
			RCD.MdiParent = this;
			RCD.Show();
		}

		private void mnuLotteReporteIPH_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteIngresoXHora RIPH = new frmReporteIngresoXHora();
			RIPH.MdiParent = this;
			RIPH.Show();
		}

		private void mnuRolReportesAE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteRolPrestamo RRP = new frmReporteRolPrestamo();
			RRP.MdiParent = this;
			RRP.Show();
		}

		private void mnuContabilidadFlujoPMP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frm_consulta_planificacion_pagos CPP = new frm_consulta_planificacion_pagos();
			CPP.MdiParent = this;
			CPP.Show();
		}

		private void mnuCrediPOINTRMR_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmMercaderiaRetiroPorLocales MRL = new frmMercaderiaRetiroPorLocales();
			MRL.MdiParent = this;
			MRL.Show();
		}

		private void mnuCrediPOINTReportesRMR_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteMRL RMR = new frmReporteMRL();
			RMR.MdiParent = this;
			RMR.Show();
		}

		private void mnuVentasReportesCXC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteCXC RCXC = new frmReporteCXC();
			RCXC.MdiParent = this;
			RCXC.Show();
		}

		private void mnuVentasPFR_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPrefacturaRemate PFR = new frmPrefacturaRemate();
			PFR.MdiParent = this;
			PFR.Show();
		}

		private void mnuCPReportesEC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmEstadoDeCartera REC = new frmEstadoDeCartera();
			REC.MdiParent = this;
			REC.Show();
		}

		private void mnuPSRDC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmResumenCierreCaja RCC = new frmResumenCierreCaja();
			RCC.MdiParent = this;
			RCC.Show();
		}

		private void mnuRolesReportesHHT_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteHE HHT = new frmReporteHE();
			HHT.MdiParent = this;
			HHT.Show();
		}

		private void mnuCPCobranzasCCPPC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCCPPC CCPPC = new frmCCPPC();
			CCPPC.MdiParent = this;
			CCPPC.Show();
		}

		private void mnuCPReportesREC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
//			frmCre_ReportesCartera CPC = new frmCre_ReportesCartera();
//			CPC.MdiParent = this;
//			CPC.Show();

			frmCre_EstadoCartera CEC = new frmCre_EstadoCartera();
			CEC.MdiParent = this;
			CEC.Show();
		}

		private void mnuVentasReportesVPA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteVendedorPorAño VPA = new frmReporteVendedorPorAño();
			VPA.MdiParent = this;
			VPA.Show();
		}

		private void mnuPCGEx_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPoliticas Politicas = new frmPoliticas();
			Politicas.MdiParent = this;
			Politicas.Show();
		}

		private void mnuVentasRP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteRecepcionDeProductos RRP = new frmReporteRecepcionDeProductos();
			RRP.MdiParent = this;
			RRP.Show();
		}

		private void mnuVentasECW_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteCuponesWeb RCW = new frmReporteCuponesWeb();
			RCW.MdiParent = this;
			RCW.Show();
		}

		private void mnuHerrVerifDepositos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmBuscarDeposito BD = new frmBuscarDeposito(false);
			BD.MdiParent = this;
			BD.Show();
		}

		private void mnuHerrSegClientes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmSeguimientoWEB SW = new frmSeguimientoWEB();
			SW.MdiParent = this;
			SW.Show();
		}

		private void mnuHerrDepositos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmDepositos VD = new frmDepositos();
			VD.MdiParent = this;
			VD.Show();
		}

		private void mnuContabilidadFPPP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmProyeccionDePagosQuincenal PPQ = new frmProyeccionDePagosQuincenal();
			PPQ.MdiParent = this;
			PPQ.Show();
		}

		private void mnuVentasReportesVentasPorArticulos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteVentaPorArticulo RVPAS = new frmReporteVentaPorArticulo();
			RVPAS.MdiParent = this;
			RVPAS.Show();
		}

		private void mnuVentasReportesVentasarticuloPorMes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteArticulosVentaPorMes RVPASM = new frmReporteArticulosVentaPorMes ();
			RVPASM.MdiParent = this;
			RVPASM.Show ();
		}

		private void mnuComprasRepCodigoArticulos_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteCodigoArticulo FRMRCOD = new frmReporteCodigoArticulo ();
			FRMRCOD.MdiParent = this;
			FRMRCOD.Show ();
		}

		private void mnuDTSubGrupo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			ReporteDTSubGrupo FRMDTSGP  = new ReporteDTSubGrupo();
			FRMDTSGP.MdiParent = this;
			FRMDTSGP.Show ();

		}

		private void mnuComprasCatalogo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCatalogo1 C = new frmCatalogo1();
			C.MdiParent = this;
			C.Show ();
		}

		private void mnuBodegaRepMD_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmVerificaKardexInventario VKI = new frmVerificaKardexInventario();
			VKI.MdiParent = this;
			VKI.Show();
		}

		private void mnuCPCobranzasEDC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ReportesEstadisticosCartera FRMRESCAT = new frmCre_ReportesEstadisticosCartera();
			//			frmCre_EstadoCartera FRMRESCAT = new frmCre_EstadoCartera ();
			FRMRESCAT.MdiParent = this;
			FRMRESCAT.Show ();
		}

		private void mnuContRB_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmImportaRetencionesBancarias IRB = new frmImportaRetencionesBancarias();
			IRB.MdiParent = this;
			IRB.Show();
		}

		private void mnuBodegaTRM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmTurnosMercaderia TIM = new frmTurnosMercaderia();
			TIM.MdiParent = this;
			TIM.Show();
		}

		private void mnuVentasRPA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCobrosExternos CE = new frmCobrosExternos();
			CE.MdiParent = this;
			CE.Show();
		}

		private void mnuRolesComisionesPS_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_PresupuestoSemanal CPS = new frmCom_PresupuestoSemanal();
			CPS.MdiParent = this;
			CPS.Show();
		}

		private void barButtonItem53_ItemClick_2(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_ResumenMensual RCS = new frmCom_ResumenMensual();
			RCS.MdiParent = this;
			RCS.Show();
		}

		private void mnuTarjetasDeIngreso_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmTarjetasIngreso TDI = new frmTarjetasIngreso();
			TDI.MdiParent = this;
			TDI.Show();
		}

		private void mnuVentasCNOS_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteNumeroOperacion RNO = new frmReporteNumeroOperacion();
			RNO.MdiParent = this;
			RNO.Show();
		}

		private void mnuDirectorio_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmLocalesMatriz LM= new frmLocalesMatriz();
			LM.MdiParent = this;
			LM.Show();
		}

		private void mnuCPReportesRRM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ResumenRetirosDeMercaderia RRM = new frmCre_ResumenRetirosDeMercaderia();
			RRM.MdiParent = this;
			RRM.Show();
		}

		private void mnuComprasProformas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmProformas Prfm = new frmProformas ();
			Prfm.MdiParent = this;
			Prfm.Show ();
		}

		private void mnuRolReportesIND_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRolReportesNominaIngresos RNI = new frmRolReportesNominaIngresos();
			RNI.MdiParent = this;
			RNI.Show();
		}

		private void mnuCPReportesCVPM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_ReporteCarteraVencidaPorMes RCVM = new frmCre_ReporteCarteraVencidaPorMes();
			RCVM.MdiParent = this;
			RCVM.Show();
		}

//		private void mnuVentasReportesIngresoDeCaja_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
//		{
//			frmReporteDeIngresoCaja RIDCd = new frmReporteDeIngresoCaja();
//			RIDCd.MdiParent = this;
//			RIDCd.Show();
//		}

		private void mnuCPCobranzasMRPR_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_RetirosReservados RR = new frmCre_RetirosReservados();
			RR.MdiParent = this;
			RR.Show();
		}

		private void mnuHerramientasPTA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPlanificacionActividadesLocales PTA = new frmPlanificacionActividadesLocales();
			PTA.MdiParent = this;
			PTA.Show();
		}

		private void mnuCPCobranzasGDC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCobranzasGestor CG = new frmCobranzasGestor();
			CG.MdiParent = this;
			CG.Show();
		}

//		private void mnuVentasReportesEstadosDeCajas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
//		{
//			frmReporteDiarioCaja RPDC = new frmReporteDiarioCaja ();
//			RPDC.MdiParent = this;
//			RPDC.Show();
//
//		}

		private void mnuVentaRMUA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteMargenUtilidadPorArticulo RMUA = new frmReporteMargenUtilidadPorArticulo ();
			RMUA.MdiParent = this;
			RMUA.Show();
		}

		private void mnuContRepEstadoCB_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteEstadosConciliacionBancaria RESTCB = new frmReporteEstadosConciliacionBancaria ();
			RESTCB.MdiParent = this;
			RESTCB.Show();
		
		}

		private void mnuContEstadosCajas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteEstadoDeCajas EstadosCajas = new frmReporteEstadoDeCajas();
			EstadosCajas.MdiParent = this;
			EstadosCajas.Show();
		}

		private void mnuContIngresoDiarioCaja_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteIngresosCajaDiarios RICD = new frmReporteIngresosCajaDiarios();
			RICD.MdiParent = this;
			RICD.Show();
		}

		private void mnuContIngresosDeCajas_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDeIngresoCaja RIDCd = new frmReporteDeIngresoCaja();
			RIDCd.MdiParent = this;
			RIDCd.Show();
		}

		private void mnuContEstadoGeneralDeCaja_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDiarioCaja RPDC = new frmReporteDiarioCaja ();
			RPDC.MdiParent = this;
			RPDC.Show();

		}

		private void mnuContRepPresupAlm_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
		
			frmProyeccionPresupuestoAlmacen RPDCA = new frmProyeccionPresupuestoAlmacen ();
			RPDCA.MdiParent = this;
			RPDCA.Show();
		}

		private void mnuHerramientasAPL_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAsignaPersonalLocales APRL = new frmAsignaPersonalLocales  ();
      APRL.MdiParent = this;
      APRL.Show();
		}

		private void mnuHerramientasASDP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{

			frmRequerimientosLocales APLRL = new frmRequerimientosLocales  ();
			APLRL.MdiParent = this;
			APLRL.Show();
		}

		private void mnuCPRevision_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{			
//			frmCred_RevisiondeCedulas RC = new frmCred_RevisiondeCedulas();
//			RC.MdiParent = this;
//			RC.Show();
		}

		private void mnuCrediPointSolicitud_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
		
		}

		private void mnuCPCobranzasCarteraAIP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCobranzaAsignacionCartera CAC = new frmCobranzaAsignacionCartera();
			CAC.MdiParent = this;
			CAC.Show();
		}

		private void mnuCPCobranzasRPE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteRefinanciaciones RR = new frmReporteRefinanciaciones();
			RR.MdiParent = this;
			RR.Show();
		}

		private void mnuRolesDecimoTercerSueldo_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteRolesAnuales RRADT = new frmReporteRolesAnuales();
			RRADT.MdiParent = this;
			RRADT.Show();
		}

		private void mnuCPCDRDC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_VistaRevisiondeCedulas CVRDC = new frmCre_VistaRevisiondeCedulas ();
			CVRDC.MdiParent = this;
			CVRDC.Show();
		}

		private void mnuCPCDRVC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCred_RevisiondeCedulas CDRVC  = new frmCred_RevisiondeCedulas ("");
			CDRVC.MdiParent = this;
			CDRVC.Show ();
		}

		private void mnuCPCDSR_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCred_SolicitudesCredito CDSLCD = new frmCred_SolicitudesCredito ();
			CDSLCD.MdiParent = this;
			CDSLCD.Show ();
		}

		private void mnuCPCDRSLT_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_Solicitudes CREDSOL = new frmCre_Solicitudes ();
			CREDSOL.MdiParent = this;
			CREDSOL.Show ();
		}

		private void mnuCPCDLCPG_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCred_ReporteVerificacionTelefonica  RTVT = new frmCred_ReporteVerificacionTelefonica ();
			RTVT.MdiParent = this;
			RTVT.Show ();
		}

		private void mnuBodegaReportesKardexCosto_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmMovimientosKardexCosto KC = new frmMovimientosKardexCosto();
			KC.MdiParent = this;
			KC.Show ();
		}

		private void mnuIngreso_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmControlDeAsistencia CrtAs = new frmControlDeAsistencia();
			CrtAs.MdiParent = this;
			CrtAs.Show ();
		}

		private void mnuRPTVentasPorcMargUtilidad_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReportePorcentajesMargenUtilidadLocal RPORMARG = new frmReportePorcentajesMargenUtilidadLocal();
			RPORMARG.MdiParent = this;
			RPORMARG.Show ();
		}

		private void mnuRPtBodegaInvFEcha_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmInventarioPorFecha IFP = new frmInventarioPorFecha();
			IFP.MdiParent = this;
			IFP.Show();
		}

		private void mnuRPtBodegaRptTrans_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDeTransferencias RDTt = new frmReporteDeTransferencias();
			RDTt.MdiParent = this;
			RDTt.Show();
		}

		private void mnuRPtBodegaInvVal_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmInventarioCostoPorBodega CostoPB = new frmInventarioCostoPorBodega();
			CostoPB.MdiParent = this;
			CostoPB.Show();
		}

		private void mnuRPtBodegaKXSer_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmKardexSeriales KSs = new frmKardexSeriales();
			KSs.MdiParent = this;
			KSs.Show();
		}

		private void mnuRPtBodegaEsConInv_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteConteoDiarioInventario RCDIP = new frmReporteConteoDiarioInventario();
			RCDIP.MdiParent = this;
			RCDIP.Show();
		}

		private void mnuRPtBodegaMovDesc_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmVerificaKardexInventario VKIVK = new frmVerificaKardexInventario();
			VKIVK.MdiParent = this;
			VKIVK.Show();
		}

		private void mnuRPtBodegaKCC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmMovimientosKardexCosto KCC = new frmMovimientosKardexCosto();
			KCC.MdiParent = this;
			KCC.Show ();
		}

		private void mnuRPtContabilidadCont_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			frmReportesDeContabilidad Reporte =  new frmReportesDeContabilidad();	
			Reporte.MdiParent = this;
			Reporte.Show();

			frmInformesDeContabilidad IDC = new frmInformesDeContabilidad();
			IDC.MdiParent = this;
			IDC.Show();
			Cursor = Cursors.Default;
		}

		private void mnuRPtContabilidadRet_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			//			ResumenRetencion miRetencion = new ResumenRetencion();
			//			miRetencion.MdiParent = this;
			//			miRetencion.Show();

			frmReporteRetenciones Reporte = new frmReporteRetenciones();
			Reporte.MdiParent = this;
			Reporte.Show();
			Cursor = Cursors.Default;
		}

		private void mnuRPtContabilidadAsDEsc_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAsientosDescuadrados AsientosDesc = new frmAsientosDescuadrados();
			AsientosDesc.MdiParent = this;
			AsientosDesc.Show();
		}

		private void mnuRPtContabilidadGast_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDeGastos CRG = new frmReporteDeGastos();
			CRG.MdiParent = this;
			CRG.Show();
		}

		private void mnuRPtContabilidadFactEl_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRepFElectronica RFE = new frmRepFElectronica();
			RFE.MdiParent = this;
			RFE.Show();
		}

		private void mnuRPtContabilidadCaCh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteCajasChicas RCC = new frmReporteCajasChicas();
			RCC.MdiParent = this;
			RCC.Show();
		}

		private void mnuRPtContabilidadGaAn_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDeGastosAnual RGA = new frmReporteDeGastosAnual();
			RGA.MdiParent = this;
			RGA.Show();
		}

		private void mnuRPtContabilidadEstaConB_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteEstadosConciliacionBancaria RESTCB = new frmReporteEstadosConciliacionBancaria ();
			RESTCB.MdiParent = this;
			RESTCB.Show();
		
		}

		private void mnuRPtContabilidadPrXAlm_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmProyeccionPresupuestoAlmacen RPDCA = new frmProyeccionPresupuestoAlmacen ();
			RPDCA.MdiParent = this;
			RPDCA.Show();
		}

		private void mnuRPtRolesPFA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmNominaReporteSalidas NRS = new frmNominaReporteSalidas();
			NRS.MdiParent = this;
			NRS.Show();
		}

		private void mnuRPtRolesVeXM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCom_ReporteVendedores RV = new frmCom_ReporteVendedores();
			RV.MdiParent = this;
			RV.Show();
		}

		private void mnuRPtRolesUni_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmNominaReportes NRU = new frmNominaReportes();
			NRU.MdiParent = this;
			NRU.Show();
		}

		private void mnuRPtRolesNG_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteNominaGeneral RNG = new frmReporteNominaGeneral();
			RNG.MdiParent = this;
			RNG.Show();
		}

		private void mnuRPtRolesAEM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteRolPrestamo RRP = new frmReporteRolPrestamo();
			RRP.MdiParent = this;
			RRP.Show();
		}

		private void mnuRPtRolesHHT_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteHE HHT = new frmReporteHE();
			HHT.MdiParent = this;
			HHT.Show();
		}

		private void mnuRPtRolesIXD_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmRolReportesNominaIngresos RNI = new frmRolReportesNominaIngresos();
			RNI.MdiParent = this;
			RNI.Show();
		}

		private void mnuRPtComprasCPA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteDeComprasPorProducto CompraPorProducto = new frmReporteDeComprasPorProducto();
			CompraPorProducto.MdiParent = this;
			CompraPorProducto.Show();
		}

		private void mnuRPtComprasT_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmFlujoProveedores FP = new frmFlujoProveedores();
			FP.MdiParent = this;
			FP.Show();
		}

		private void mnuRPtComprasRAC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReportesArticulosPorComprar RAPC = new frmReportesArticulosPorComprar();
			RAPC.MdiParent = this;
			RAPC.Show();
		}

		private void mnuRPtComprasCA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteCodigoArticulo FRMRCOD = new frmReporteCodigoArticulo ();
			FRMRCOD.MdiParent = this;
			FRMRCOD.Show ();
		}

		private void mnuCompraAODC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteAprobacionODC APODC = new frmReporteAprobacionODC ();
			APODC.MdiParent = this;
			APODC.Show ();
		}

		private void mnuCPCobCM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteCashManagement RCM = new frmReporteCashManagement();
			RCM.MdiParent = this;
			RCM.Show();
		}

		private void mnuReporVentasPDGA_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteGastos RPTGT = new frmReporteGastos();
			RPTGT.MdiParent = this;
			RPTGT.Show();
		}

		private void mnuCRResumenCredito_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCred_ResumenCreditos RPCRC = new frmCred_ResumenCreditos ();
			RPCRC.MdiParent = this;
			RPCRC.Show();
		}

		private void barButtonItem82_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCre_GeneraLotesPichincha RPCRClp = new frmCre_GeneraLotesPichincha ();
			RPCRClp.MdiParent = this;
			RPCRClp.Show();
		}

		private void mnuCom_ReporteLM_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
		  frmCom_ListaMovimientosPorPeriodo1 RPLM = new frmCom_ListaMovimientosPorPeriodo1();
			RPLM.MdiParent = this;
			RPLM.Show();
		}

		private void mnuBodegaNI_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmNotasDeIngreso NDI = new frmNotasDeIngreso();
			NDI.MdiParent = this;
			NDI.Show();
		}

		private void mnuComprasLAC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCSListaDeProductosVendidos CSLPV = new frmCSListaDeProductosVendidos();
			CSLPV.MdiParent = this;
			CSLPV.Show();
		}

		private void mnuCPComisionesAsigCobr_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCbo_AsignacionCobradores CBO_CBr  = new frmCbo_AsignacionCobradores ();
      CBO_CBr.MdiParent = this;
			CBO_CBr.Show();
		}

		private void mnuCPComisionesTrmOpe_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCbo_AsignaPorcDias CBO_AsgDias = new frmCbo_AsignaPorcDias ();
			CBO_AsgDias.MdiParent = this;
			CBO_AsgDias.Show();
		}

		private void mnuCPCobranzasCboOpe_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmConsultaCbo_GestorDeCobranzas CBO_CLTGCOB = new frmConsultaCbo_GestorDeCobranzas ();
			CBO_CLTGCOB.MdiParent = this;
			CBO_CLTGCOB.Show();
		}

		private void mnuCPComisionesCom_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCobPorcentajeComisiones CBO_Comisiones = new frmCobPorcentajeComisiones ();
			CBO_Comisiones.MdiParent = this;
			CBO_Comisiones.Show();
		}

		private void mnuCreditoGesCall_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmConsultaGestionesCallCenter CBO_RptCall= new frmConsultaGestionesCallCenter ();
			CBO_RptCall.MdiParent = this;
			CBO_RptCall.Show();
		}

		private void mnuCreditoProReal_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCobranzasProyectadoReal  CBO_ProyeRe = new frmCobranzasProyectadoReal ();
			CBO_ProyeRe.MdiParent = this;
			CBO_ProyeRe.Show();
		}

		private void mnuRPPagoAgil_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReportePagoAgil RPPA = new frmReportePagoAgil();
			RPPA .MdiParent = this;
			RPPA .Show();
		}

		private void mnuCPCobranzasCERTIFICADO_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCred_CertificadoCobranzas CRED_CRTF  =  new frmCred_CertificadoCobranzas();
			CRED_CRTF.MdiParent = this;
			CRED_CRTF.Show();
		}

		private void mnuRPtComprasPuC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			FrmReporteVentasPorCostoUltimo Cmp_UC =  new FrmReporteVentasPorCostoUltimo();
			Cmp_UC.MdiParent = this;
			Cmp_UC.Show();
		}

		private void barButtonItem84_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{			
			frmCob_ResumenDeComisionesMensuales Cob_rC =  new frmCob_ResumenDeComisionesMensuales();
			Cob_rC.MdiParent = this;
			Cob_rC.Show();
		}

		private void mnuCPANDC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmAprobacionNotasDeCredito APNC = new frmAprobacionNotasDeCredito();
			APNC.MdiParent = this;
			APNC.Show();
		}

		private void mnuComprasPromoFact_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmFacturaPromociones FacProm = new frmFacturaPromociones();
			FacProm.MdiParent = this;
			FacProm.Show();
		}

		private void mnuSolNC_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmSolicitudNotasDeCredito SNC = new frmSolicitudNotasDeCredito();
			SNC.MdiParent = this;
			SNC.Show();
		}

		private void mnuRolPorCajero_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmPorcentajeCajeros RRPC = new frmPorcentajeCajeros();
			RRPC.MdiParent = this;
			RRPC.Show();
		}

		private void mnuRPTVentasArtXMes_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
		
		}

		private void barButtonItem86_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmProyectosZoomatica Pzoom = new frmProyectosZoomatica();
			Pzoom.MdiParent = this;
			Pzoom.Show();
		}

		private void barButtonItem87_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmConfiguracionGEXyCompraProtegida PGex = new frmConfiguracionGEXyCompraProtegida();
			PGex.MdiParent = this;
			PGex.Show();
		}

		private void mnuRPtCREPOINTFDG_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteFondoDeGarantia  RPTFDG = new frmReporteFondoDeGarantia();
			RPTFDG.MdiParent = this;
			RPTFDG.Show();
		}

		private void mnuRPtCREPOINRMD_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteRangoMorosidad fRRM = new frmReporteRangoMorosidad();
			fRRM.MdiParent = this;
			fRRM.Show();
		}

		private void mnuCPCobranzasRPP_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmCob_Operaciones_Retorno_Point RPP = new frmCob_Operaciones_Retorno_Point();
			RPP.MdiParent = this;
			RPP.Show();
		}

		private void mnuReporVentasDFI_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{			
			frmSeguroDFI FSDFI = new frmSeguroDFI();
			FSDFI.MdiParent = this;
			FSDFI.Show();			
		}

		private void mnuRPtContabilidadSCosecha_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
		{
			frmReporteSiembraCosecha RPTSCH = new frmReporteSiembraCosecha();
			RPTSCH .MdiParent = this;
			RPTSCH.Show();			
		}
	}	
}
