using System;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Acceso.
	/// </summary>
	public class Acceso
	{
		private bool imprimir =  true;
		private bool anularDocSRI = true;
		private bool buscar = true;
		private bool nuevo = true;
		private bool editar = true;
		private bool bAnulaFacturaMayor0 = true;
		private bool editarFechaEg = true;
		private bool editarSoloDia = false;
		private bool editarDiaIngreso = false;
		private bool editarCuentaNombre;
		private bool editarCuentaCodigo;
		private bool editarTabla;
		private bool anular = true;
		private bool anularCP = true;
		private bool recuperar = true;
		private bool eliminar = true;
		private bool cambiarFecha = true;
		private bool cambiarFechaDiciembre = true;
		private bool cambiarFechaPosterior = true;
		private bool cambiarPeriodo = true;
		private bool mostrarDia;
		private bool cambiarRetencion;
		private bool cambiarNumeracion;
		private bool imprimirRetencion;
		private bool imprimirCheque;
		private int nivelAcceso = 1;
		private bool crearAsiento = true;
		private bool cambiarGrupo;
		private bool transformacion = true;
		private bool retencion;
		private bool actualizaPrecio;
		private bool cambiarPrecio = true;
		private bool cambiarDescuento = true;
		private bool reImprimir = true;
		private bool editaImpresa = true;
		private bool grabaCero = true;
		private bool autorizacion = true;
		private bool listaPrecios = false;
		private bool auditoria = true;
		private bool editarRevisadas;
		private bool editarDetalleFormaPago = true;
		private bool editarContabilizado = true;
		private bool bModificarGastos = true;
		private bool bModificarMora = true;
		private bool bPrecioDistribuidor = true;
		private bool bAprobarPrefactura = true;
		private bool bRechazarPrefactura = true;
		private bool bBuscarOrdenCompra = true;
		private bool bFichaArticulos = true;
		private bool bAuditoria = true;
		private bool bDepositoCaja = true;
		private bool bAutorizar = true;
		private bool bAutorizarPedido = true;
		private bool bVistas = true;
		private bool modificarfechasanteriores = true;
		private bool bExportarExcel = true;
		private bool bCambiarVencimiento = true;
		private bool bModificarBodega = true;
		private bool bAutEntrada = true;
		private bool bAutCobro = true;
		private bool bBuscarDocumentos = true;
		private bool bActNum = true;
		private bool bEditaCobranza = true;
		private bool bAnulaCobranza = true;
		private bool bModifcarFormaDePago = true;
		private bool bContabilizarFactMan = true;
		private bool bVerProveedorDT = true;
		private bool bAPAutorizar = true;
		private bool bAPAsignaBancos = true;
		private bool bAPActAutorizados = true;
		private bool bAPVerAutorizados = true;
		private bool bAPVerPagados = true;
		private bool bSobrantesFaltantesCaja = true;
		private bool bModificaRenta = true;
		private bool bModificaIVA = true;
		private bool bEditarNombre = true;
		private bool bModArtCmbProm = true;
		private bool bEliArtCmbProm = true;
		private bool bEliArt = true;
		private bool bBloquearCB = true;
		private bool bDesbloquearCB = true;
		private bool bEditarProcesado = true;
		private bool bModArtSerializado = true;
		private bool bEliArtSerializado = true;
		private bool bAnularConSeriales = true;
		private bool bAnularProcesado = true;
		private bool bEditarAprobado = true;
		private bool bAnularAprobado = true;
		private bool bEditarPagosAutorizados = true;
		private bool bEditarPagosRealizados  = true;
		private bool bEditarLiquidado = true;		
		private bool bAnularPagosAutorizados = true;
		private bool bAnularPagosRealizados = true;
		private bool bAnularLiquidado = true;
		private bool bEliminaFormaDePago = true;
		private bool bManejaSerial = true;
		private bool bCruzaAnticiposCP = true;
		private bool bEliminaSeriales = true;
		private bool bModificaSeriales = true;
		private bool bVerSerialesProcesadas = true;
		private bool bNCFiscal = true;
		private bool bProcesarSerialesFaltantes = true;
		private bool bModificarSaldoCB = true;
		private bool bAsignaCobrador = true;
		private bool bAnulaSoloDelDia = true;
		private bool bEliminaFilaCCEnEspera = true;
		private bool bEliminaFilaCCProcesada = true;
		private bool bProcesar = true;
		private bool bAnularEnTransito = true;
		private bool bAnularTFProcesado = true;
		private bool bModificarImpreso = true;
		private bool bAbrirImpuestos = true;
		private bool bVerCapitalInteres = true;
		private bool bCarpa = true;
		private bool bExportar = true;
		private bool editarLote = true;
		private bool editarReferencia = true;
		
		private bool bTodos = true;
		private bool bDepartamentoTecnico = true;
		private bool bCompras = true;
		private bool bEstados = true;
		private bool bTransferir = true;
		private bool bModRevDT = true;
		private bool bModRevCompras = true;
		private bool bCotizaAccesorios = true;
		private bool bRevisionObsoleto = true;
	
		private bool bCom_FiltrarPorCargo = true;
		private bool bCom_CerrarPeriodo = true;
		private bool bCom_Observaciones = true;
		private bool bRolContabilizar = true;
		private bool bRolGenerar = true;
		private bool bRolRevision = true;
		private bool bRolQuitarRevisado = true;
		private bool bRolAutorizar = true;
		private bool bRolQuitarAutorizado = true;
		private bool bRolAcreditar = true;
		private bool bRolSubirArchivos = true;
		private bool bVerComisiones = true;
		private bool bGenerarCobro = true;		
		private bool bNominaEditarIngresosCargos = true;
		private bool bNominaEliminarExtras = true;
		private bool bNominaEliminarMovimientos = true;
		private bool bNominaEditarExtras = true;
		private bool bNominaEditarMovimientos = true;
		private bool bReliquidar = true;
		private bool bReliquidarCruzandoCuentas = true;
		private bool bCruzarPagosAdelantados = true;
		private bool b60DiasCobro = true;
		private bool b30DiasCobro = true;
		private bool b15DiasCobro = true;		
		private bool b1DiasCobro = true;
		private bool bDia = true;
		private bool bModificarCapitalInteres = true;
		private bool bAprobarCarpa = true;
		private bool bIrA = true;
		private bool bAsignaTecnicos = true;
		private bool bModificaGarantia = true;
		private bool bVerRemates = true;
		private bool bVerMatriz = true;
		private bool bModificarCuentas = true;
		private bool bCerrarCaja = true;
		private bool bAbrirCaja = true;
		private bool bEliminaModificaSeriales = true;
		private bool bLimiteBusquedaFechas = true;
		private bool bFormasDePagoCred = true;
		private bool bFormasDePagoCE = true;
		private bool bRevisadoCE = true;

		private bool bRefinanciar = true;
		private bool bCondonar = true;
		private bool bCobrar = true;
		private bool bHistorico = true;

		private bool bGenerarNC = true;
		private bool bGenerarNCP = true;
		private bool bArticulosAtributos = true;
		private bool bSolNCAFD = true;
		private bool bFacturar = true;
		private bool bActualizar = true;
		private bool bGenerar = true;
		private bool bImportar = true;
		private bool bVerDocumentos = true;
		private bool bAyuda = true;
		private bool bAsignar = true;
		private bool bGuardar = true;
		private bool bDescuento = true;
		private bool bAprobarValorMayor = true;
		private bool bVerMayor100dias = true;
		private bool bSupervisor = true;
		private bool bAprobado = true;
		private bool bFecha3Días = true;
		private bool bEditPulgadas = true;
		private bool bEditColor = true;
		private bool bEditProcesador = true;
		private bool bEditMarca = true;
		private bool bEditMemoria = true;
		private bool bEditDisco = true;
		private bool bEditPrecio = true;
		private bool bEditGaming = true;
		private bool bEditLumenes = true;
		private bool bEditResolucion = true;
		private bool bEditCapacidad = true;
		private bool bEditAlmacenamiento = true;
		private bool bEditPixeles = true;
		private bool bEditQuemadores = true;
		private bool bEditTarjeta = true;
		private bool bEditOfeElectro = true;
		private bool bEditOfeComp = true;
		private bool bAccesoBoton = true;
		private bool bGestion = true;
		private bool bMora = true;
		private bool bFiniquitos = true;
		private bool btnDescuento = true;
		private bool btnDistribucion = true;
		private bool bDesModifica = true;
		private bool bDesEmpleado = true;
		private bool bRefacturacion = true;
		private bool bAnulaTiempo = true;
		private bool bEditacionMinutos = true;
		private bool bCobranzasSinLimiteFechas = true;
		private bool bTipoBono = true;
		private bool bAcreditar = true;
		private bool bEditaCantidad = true;
		private bool bSubirArchivos = true;
		private bool bImagen = true;
		private bool bCupon = true;
		private bool bHabilitaControles = true;
		private bool bFacturaExtras = true;
		private bool bFTemporal = true;
		private bool bCambiarPrecioOE = true;
		private bool bAutPagosValor = true;
		private bool bCuponEmpresas = true;
		private bool bCruzaAnticiposCC = true;
		private bool bAnulaCrucesProveedores = true;
		private bool bActivaTarjetaVIP = true;
		private bool modificarMoraCartera = true;
		private bool bBuscarLista = true;
		private bool bExportaLista = true;
		private bool bHabilitaOPT = true;
		private bool abrefacturacion = true;
		private bool responsable = true;
		private bool editarresponsable = true;
		private bool cargopersonal = true;
		private bool bAnulaCancelado = true;
		private bool bAnulFacEnviada = true;
		/*965	REFINANCIAR
966	CONDONAR
967	COBRAR
968	HISTORICO*/
		public bool Auditoria
		{
			get
			{
				return auditoria;
			}
			set
			{
				auditoria = value;
			}
		}
		public bool Autorizacion
		{
			get
			{
				return autorizacion;
			}
			set
			{
				autorizacion = value;
			}
		}

		public bool ListaPrecios
		{
			get
			{
				return listaPrecios;
			}
			set
			{
				listaPrecios = value;
			}
		}

		public int NivelAcceso
		{
			get
			{
				return nivelAcceso;
			}
			set
			{
				nivelAcceso = value;
			}
		}

		public bool GrabaCero
		{
			get
			{
				return grabaCero;
			}
			set
			{
				grabaCero = value;
			}
		}

		public Acceso()
		{
			
		}

		public bool Buscar
		{
			get
			{
				return buscar;
			}
			set 
			{
				buscar = value;
			}
		}
		public bool Imprimir
		{
			get
			{
				return imprimir;
			}
			set 
			{
				imprimir = value;
			}
		}
		public bool AnularDocSRI
		{
			get
			{
				return anularDocSRI;
			}
			set 
			{
				anularDocSRI = value;
			}
		}
		public bool ReImprimir
		{
			get
			{
				return reImprimir;
			}
			set 
			{
				reImprimir = value;
			}
		}
		public bool Nuevo
		{
			get
			{
				return nuevo;
			}
			set 
			{
				nuevo = value;
			}
		}
		public bool Editar
		{
			get
			{
				return editar;
			}
			set 
			{
				editar = value;
			}
		}
		public bool BAnulaFacturaMayor0
		{
			get
			{
				return bAnulaFacturaMayor0;
			}
			set 
			{
				bAnulaFacturaMayor0 = value;
			}
		}

		public bool EditarFechaEg
		{
			get
			{
				return editarFechaEg;
			}
			set 
			{
				editarFechaEg = value;
			}
		}

		public bool CambiarFechaPosterior
		{
			get
			{
				return cambiarFechaPosterior;
			}
			set 
			{
				cambiarFechaPosterior = value;
			}
		}

		public bool EditarTabla
		{
			get
			{
				return editarTabla;
			}
			set 
			{
				editarTabla = value;
			}
		}
		public bool EditarCuentaNombre
		{
			get
			{
				return editarCuentaNombre;
			}
			set 
			{
				editarCuentaNombre = value;
			}
		}
		public bool EditarCuentaCodigo
		{
			get
			{
				return editarCuentaCodigo;
			}
			set 
			{
				editarCuentaCodigo = value;
			}
		}

		public bool EditarSoloDia
		{
			get
			{
				return editarSoloDia;
			}
			set 
			{
				editarSoloDia = value;
			}
		}
		public bool EditarDiaIngreso
		{
			get
			{
				return editarDiaIngreso;
			}
			set 
			{
				editarDiaIngreso = value;
			}
		}
		public bool Anular
		{
			get
			{
				return anular;
			}
			set 
			{
				anular = value;
			}
		}
		public bool AnularCrediPoint
		{
			get
			{
				return anularCP;
			}
			set 
			{
				anularCP = value;
			}
		}
		
		public bool CrearAsiento
		{
			get
			{
				return crearAsiento;
			}
			set 
			{
				crearAsiento = value;
			}
		}
		public bool Recuperar
		{
			get
			{
				return recuperar;
			}
			set 
			{
				recuperar = value;
			}
		}
		public bool Eliminar
		{
			get
			{
				return eliminar;
			}
			set 
			{
				eliminar = value;
			}
		}
		public bool CambiarFecha
		{
			get
			{
				return cambiarFecha;
			}
			set 
			{
				cambiarFecha = value;
			}
		}
		public bool CambiarFechaDiciembre
		{
			get
			{
				return cambiarFechaDiciembre;
			}
			set 
			{
				cambiarFechaDiciembre = value;
			}
		}
		public bool CambiarPeriodo
		{
			get
			{
				return cambiarPeriodo;
			}
			set 
			{
				cambiarPeriodo = value;
			}
		}

		public bool MostrarDia
		{
			get
			{
				return mostrarDia;
			}
			set 
			{
				mostrarDia = value;
			}
		}
		
		public bool CambiarNumeracion
		{
			get
			{
				return cambiarNumeracion;
			}
			set 
			{
				cambiarNumeracion = value;
			}
		}
		public bool CambiarRetencion
		{
			get
			{
				return cambiarRetencion;
			}
			set 
			{
				cambiarRetencion = value;
			}
		}
		public bool ImprimirRetencion
		{
			get
			{
				return imprimirRetencion;
			}
			set 
			{
				imprimirRetencion = value;
			}
		}
		public bool ImprimirCheque
		{
			get
			{
				return imprimirCheque;
			}
			set 
			{
				imprimirCheque = value;
			}
		}

		public bool CambiarGrupo
		{
			get
			{
				return cambiarGrupo;
			}
			set 
			{
				cambiarGrupo = value;
			}
		}		
		
		public bool Transformacion
		{
			get
			{
				return transformacion;
			}
			set 
			{
				transformacion = value;
			}
		}

		public bool Retencion
		{
			get
			{
				return retencion;
			}
			set 
			{
				retencion = value;
			}
		}
		public bool ActualizaPrecio
		{
			get
			{
				return actualizaPrecio;
			}
			set 
			{
				actualizaPrecio = value;
			}
		}
		public bool CambiarPrecio
		{
			get
			{
				return cambiarPrecio;
			}
			set 
			{
				cambiarPrecio = value;
			}
		}
		public bool CambiarDescuento
		{
			get
			{
				return cambiarDescuento;
			}
			set 
			{
				cambiarDescuento = value;
			}
		}

		public bool EditaImpresa
		{
			get
			{
				return editaImpresa;
			}
			set 
			{
				editaImpresa = value;
			}
		}
		public bool EditarRevisadas
		{
			get
			{
				return editarRevisadas;
			}
			set 
			{
				editarRevisadas = value;
			}
		}
		
		public bool EditarDetalleFormaPago
		{
			get
			{
				return editarDetalleFormaPago;
			}
			set 
			{
				editarDetalleFormaPago = value;
			}
		}

		public bool EditarContabilizado
		{
			get
			{
				return editarContabilizado;
			}
			set 
			{
				editarContabilizado = value;
			}
		}

		public bool BModificarMora
		{
			get
			{
				return bModificarMora;
			}
			set 
			{
				bModificarMora = value;
			}
		}

		public bool BModificarGastos
		{
			get
			{
				return bModificarGastos;
			}
			set 
			{
				bModificarGastos = value;
			}
		}

		public bool BPrecioDistribuidor
		{
			get
			{
				return bPrecioDistribuidor;
			}
			set 
			{
				bPrecioDistribuidor = value;
			}
		}

		public bool BAprobarPrefactura
		{
			get
			{
				return bAprobarPrefactura;
			}
			set 
			{
				bAprobarPrefactura = value;
			}
		}

		public bool BRechazarPrefactura
		{
			get
			{
				return bRechazarPrefactura;
			}
			set 
			{
				bRechazarPrefactura = value;
			}
		}
		
		public bool BBuscarOrdenCompra
		{
			get
			{
				return bBuscarOrdenCompra;
			}
			set 
			{
				bBuscarOrdenCompra = value;
			}
		}

		public bool BFichaArticulos
		{
			get
			{
				return bFichaArticulos;
			}
			set 
			{
				bFichaArticulos = value;
			}
		}

		public bool BAuditoria
		{
			get
			{
				return bAuditoria;
			}
			set 
			{
				bAuditoria = value;
			}
		}

		public bool BDepositoCaja
		{
			get
			{
				return bDepositoCaja;
			}
			set 
			{
				bDepositoCaja = value;
			}
		}
		
		public bool BAutorizar
		{
			get
			{
				return bAutorizar;
			}
			set
			{
				bAutorizar = value;
			}
		}

		public bool BAutorizarPedido
		{
			get
			{
				return bAutorizarPedido;
			}
			set
			{
				bAutorizarPedido = value;
			}
		}

		public bool BVistas
		{
			get
			{
				return bVistas;
			}
			set
			{
				bVistas = value;
			}
		}

		public bool BExportarExcel
		{
			get
			{
				return bExportarExcel;
			}
			set
			{
				bExportarExcel = value;
			}
		}

		public bool BCambiarVencimiento
		{
			get
			{
				return bCambiarVencimiento;
			}
			set
			{
				bCambiarVencimiento = value;
			}
		}

		public bool BModificarBodega
		{
			get
			{
				return bModificarBodega;
			}
			set
			{
				bModificarBodega = value;
			}
		}

		public bool BAutEntrada
		{
			get
			{
				return bAutEntrada;
			}
			set
			{
				bAutEntrada = value;
			}
		}

		public bool BAutCobro
		{
			get
			{
				return bAutCobro;
			}
			set
			{
				bAutCobro = value;
			}
		}

		public bool BBuscarDocumentos
		{
			get
			{
				return bBuscarDocumentos;
			}
			set
			{
				bBuscarDocumentos = value;
			}
		}

		public bool BActNum 
		{
			get
			{
				return bActNum;
			}
			set
			{
				bActNum = value;
			}
		}

		public bool BEditaCobranza
		{
			get
			{
				return bEditaCobranza;
			}
			set
			{
				bEditaCobranza = value;
			}
		}

		public bool BAnulaCobranza 
		{
			get
			{
				return bAnulaCobranza;
			}
			set
			{
				bAnulaCobranza = value;
			}
		}

		public bool BModifcarFormaDePago 
		{
			get
			{
				return bModifcarFormaDePago;
			}
			set
			{
				bModifcarFormaDePago = value;
			}
		}
		
		public bool BContabilizarFactMan 
		{
			get
			{
				return bContabilizarFactMan;
			}
			set
			{
				bContabilizarFactMan = value;
			}
		}

		public bool BVerProveedorDT
		{
			get
			{
				return bVerProveedorDT;
			}
			set
			{
				bVerProveedorDT = value;
			}
		}

		public bool BSobrantesFaltantesCaja
		{
			get
			{
				return bSobrantesFaltantesCaja;
			}
			set
			{
				bSobrantesFaltantesCaja = value;
			}
		}

		#region Servicios 
		public bool BModificaRenta
		{
			get
			{
				return bModificaRenta;
			}
			set
			{
				bModificaRenta = value;
			}
		}

		public bool BModificaIVA
		{
			get
			{
				return bModificaIVA;
			}
			set
			{
				bModificaIVA = value;
			}
		}
		#endregion Servicios 

		#region Autorizacion de Pagos
		public bool BAPAutorizar 
		{
			get
			{
				return bAPAutorizar;
			}
			set
			{
				bAPAutorizar = value;
			}
		}

		public bool BAPAsignaBancos
		{
			get
			{
				return bAPAsignaBancos;
			}
			set
			{
				bAPAsignaBancos = value;
			}
		}

		public bool BAPActAutorizados
		{
			get
			{
				return bAPActAutorizados;
			}
			set
			{
				bAPActAutorizados = value;
			}
		}

		public bool BAPVerAutorizados
		{
			get
			{
				return bAPVerAutorizados;
			}
			set
			{
				bAPVerAutorizados = value;
			}
		}

		public bool BAPVerPagados
		{
			get
			{
				return bAPVerPagados;
			}
			set
			{
				bAPVerPagados = value;
			}
		}		
		#endregion Autorizacion de Pagos
			
		public bool BEditarNombre
		{
			get
			{
				return bEditarNombre;
			}
			set
			{
				bEditarNombre = value;
			}
		}

		public bool BModArtCmbProm
		{
			get
			{
				return bModArtCmbProm;
			}
			set
			{
				bModArtCmbProm = value;
			}
		}

		public bool BEliArtCmbProm
		{
			get
			{
				return bEliArtCmbProm;
			}
			set
			{
				bEliArtCmbProm = value;
			}
		}

		public bool BEliArt
		{
			get
			{
				return bEliArt;
			}
			set
			{
				bEliArt = value;
			}
		}
		
		public bool BBloquearCB
		{
			get
			{
				return bBloquearCB;
			}
			set
			{
				bBloquearCB = value;
			}
		}

		public bool BDesbloquearCB
		{
			get
			{
				return bDesbloquearCB;
			}
			set
			{
				bDesbloquearCB = value;
			}
		}

		public bool BEditarProcesado
		{
			get
			{
				return bEditarProcesado;
			}
			set
			{
				bEditarProcesado = value;
			}
		}
		
		public bool BModArtSerializado
		{
			get
			{
				return bModArtSerializado;
			}
			set
			{
				bModArtSerializado = value;
			}
		}

		public bool BEliArtSerializado
		{
			get
			{
				return bEliArtSerializado;
			}
			set
			{
				bEliArtSerializado = value;
			}
		}

		public bool BAnularConSeriales
		{
			get
			{
				return bAnularConSeriales;
			}
			set
			{
				bAnularConSeriales = value;
			}
		}

		public bool BAnularProcesado
		{
			get
			{
				return bAnularProcesado;
			}
			set
			{
				bAnularProcesado = value;
			}
		}
		
		public bool BEditarAprobado
		{
			get
			{
				return bEditarAprobado;
			}
			set
			{
				bEditarAprobado = value;
			}
		}

		public bool BAnularAprobado
		{
			get
			{
				return bAnularAprobado;
			}
			set
			{
				bAnularAprobado = value;
			}
		}
		
		public bool BEditarPagosAutorizados
		{
			get
			{
				return bEditarPagosAutorizados;
			}
			set
			{
				bEditarPagosAutorizados = value;
			}
		}

		public bool BEditarPagosRealizados
		{
			get
			{
				return bEditarPagosRealizados;
			}
			set
			{
				bEditarPagosRealizados = value;
			}
		}

		public bool BEditarLiquidado
		{
			get
			{
				return bEditarLiquidado;
			}
			set
			{
				bEditarLiquidado = value;
			}
		}

		public bool BAnularPagosAutorizados
		{
			get
			{
				return bAnularPagosAutorizados;
			}
			set
			{
				bAnularPagosAutorizados = value;
			}
		}
		
		public bool BAnularPagosRealizados
		{
			get
			{
				return bAnularPagosRealizados;
			}
			set
			{
				bAnularPagosRealizados = value;
			}
		}

		public bool BAnularLiquidado
		{
			get
			{
				return bAnularLiquidado;
			}
			set
			{
				bAnularLiquidado = value;
			}
		}

		public bool BEliminaFormaDePago
		{
			get
			{
				return bEliminaFormaDePago;
			}
			set
			{
				bEliminaFormaDePago = value;
			}
		}

		public bool BManejaSerial
		{
			get
			{
				return bManejaSerial;
			}
			set
			{
				bManejaSerial = value;
			}
		}

		public bool BCruzaAnticiposCP
		{
			get
			{
				return bCruzaAnticiposCP;
			}
			set
			{
				bCruzaAnticiposCP = value;
			}
		}
		
		public bool BEliminaSeriales
		{
			get
			{
				return bEliminaSeriales;
			}
			set
			{
				bEliminaSeriales = value;
			}
		}

		public bool BModificaSeriales
		{
			get
			{
				return bModificaSeriales;
			}
			set
			{
				bModificaSeriales = value;
			}
		}

		public bool BVerSerialesProcesadas
		{
			get
			{
				return bVerSerialesProcesadas;
			}
			set
			{
				bVerSerialesProcesadas = value;
			}
		}
		
		public bool BNCFiscal
		{
			get
			{
				return bNCFiscal;
			}
			set
			{
				bNCFiscal = value;
			}
		}

		public bool BProcesarSerialesFaltantes
		{
			get
			{
				return bProcesarSerialesFaltantes;
			}
			set
			{
				bProcesarSerialesFaltantes = value;
			}
		}

		public bool BModificarSaldoCB
		{
			get
			{
				return bModificarSaldoCB;
			}
			set
			{
				bModificarSaldoCB = value;
			}
		}

		public bool BAsignaCobrador
		{
			get
			{
				return bAsignaCobrador;
			}
			set
			{
				bAsignaCobrador = value;
			}
		}
		
		public bool BAnulaSoloDelDia 
		{
			get
			{
				return bAnulaSoloDelDia;
			}
			set
			{
				bAnulaSoloDelDia = value;
			}
		}
		
		public bool BEliminaFilaCCEnEspera
		{
			get
			{
				return bEliminaFilaCCEnEspera ;
			}
			set
			{
				bEliminaFilaCCEnEspera = value;
			}
		}
		
		public bool BEliminaFilaCCProcesada
		{
			get
			{
				return bEliminaFilaCCProcesada;
			}
			set
			{
				bEliminaFilaCCProcesada = value;
			}
		}

		public bool BProcesar
		{
			get
			{
				return bProcesar;
			}
			set
			{
				bProcesar = value;
			}
		}

		
		
		public bool BAnularEnTransito
		{
			get
			{
				return bAnularEnTransito;
			}
			set
			{
				bAnularEnTransito = value;
			}
		}

		public bool BAnularTFProcesado
		{
			get
			{
				return bAnularTFProcesado;
			}
			set
			{
				bAnularTFProcesado = value;
			}
		}

		public bool BModificarImpreso
		{
			get
			{
				return bModificarImpreso;
			}
			set
			{
				bModificarImpreso = value;
			}
		}

		public bool BAbrirImpuestos
		{
			get
			{
				return bAbrirImpuestos;
			}
			set
			{
				bAbrirImpuestos = value;
			}
		}

		public bool BVerCapitalInteres
		{
			get
			{
				return bVerCapitalInteres;
			}
			set
			{
				bVerCapitalInteres = value;
			}
		}
		
		public bool BCarpa
		{
			get
			{
				return bCarpa;
			}
			set
			{
				bCarpa = value;
			}
		}

		public bool BExportar
		{
			get
			{
				return bExportar;
			}
			set
			{
				bExportar = value;
			}
		}

		#region Informe Mercaderia de remates		
		public bool BTodos
		{
			get
			{
				return bTodos;
			}
			set
			{
				bTodos = value;
			}
		}

		public bool BDepartamentoTecnico
		{
			get
			{
				return bDepartamentoTecnico;
			}
			set
			{
				bDepartamentoTecnico = value;
			}
		}

		public bool BCompras
		{
			get
			{
				return bCompras;
			}
			set
			{
				bCompras = value;
			}
		}

		public bool BEstados
		{
			get
			{
				return bEstados;
			}
			set
			{
				bEstados = value;
			}
		}

		public bool BTransferir
		{
			get
			{
				return bTransferir;
			}
			set
			{
				bTransferir = value;
			}
		}
		
		public bool BModRevDT
		{
			get
			{
				return bModRevDT;
			}
			set
			{
				bModRevDT = value;
			}
		}

		public bool BModRevCompras
		{
			get
			{
				return bModRevCompras;
			}
			set
			{
				bModRevCompras = value;
			}
		}

		public bool BCotizaAccesorios
		{
			get
			{
				return bCotizaAccesorios;
			}
			set
			{
				bCotizaAccesorios = value;
			}
		}

		public bool BRevisionObsoleto
		{
			get
			{
				return bRevisionObsoleto;
			}
			set
			{
				bRevisionObsoleto = value;
			}
		}
		#endregion Informe Mercaderia de remates

		public bool BCom_CerrarPeriodo
		{
			get
			{
				return bCom_CerrarPeriodo;
			}
			set
			{
				bCom_CerrarPeriodo = value;
			}
		}

		public bool BCom_FiltrarPorCargo
		{
			get
			{
				return bCom_FiltrarPorCargo;
			}
			set
			{
				bCom_FiltrarPorCargo = value;
			}
		}
		
		public bool BCom_Observaciones
		{
			get
			{
				return bCom_Observaciones;
			}
			set
			{
				bCom_Observaciones = value;
			}
		}
		
		public bool BVerComisiones
		{
			get
			{
				return bVerComisiones;
			}
			set
			{
				bVerComisiones = value;
			}
		}

		#region Roles de Pago
		public bool BRolContabilizar 
		{
			get
			{
				return bRolContabilizar ;
			}
			set
			{
				bRolContabilizar = value;
			}
		}

		public bool BRolGenerar 
		{
			get
			{
				return bRolGenerar ;
			}
			set
			{
				bRolGenerar = value;
			}
		}

		public bool BRolRevision 
		{
			get
			{
				return bRolRevision ;
			}
			set
			{
				bRolRevision = value;
			}
		}

		public bool BRolQuitarRevisado 
		{
			get
			{
				return bRolQuitarRevisado ;
			}
			set
			{
				bRolQuitarRevisado = value;
			}
		}
		
		public bool BRolAutorizar 
		{
			get
			{
				return bRolAutorizar;
			}
			set
			{
				bRolAutorizar = value;
			}
		}

		public bool BRolQuitarAutorizado  
		{
			get
			{
				return bRolQuitarAutorizado  ;
			}
			set
			{
				bRolQuitarAutorizado = value;
			}
		}
		
		public bool BRolAcreditar 
		{
			get
			{
				return bRolAcreditar;
			}
			set
			{
				bRolAcreditar = value;
			}
		}

		public bool BRolSubirArchivos
		{
			get
			{
				return bRolSubirArchivos;
			}
			set
			{
				bRolSubirArchivos = value;
			}
		}
		#endregion Roles de Pago

		public bool BGenerarCobro
		{
			get
			{
				return bGenerarCobro;
			}
			set
			{
				bGenerarCobro = value;
			}
		}

		public bool BAprobarCarpa
		{
			get
			{
				return bAprobarCarpa;
			}
			set
			{
				bAprobarCarpa = value;
			}
		}

		public bool BIrA
		{
			get
			{
				return bIrA;
			}
			set
			{
				bIrA = value;
			}
		}

		#region Nomina
		public bool BNominaEditarIngresosCargos
		{
			get
			{
				return bNominaEditarIngresosCargos;
			}
			set
			{
				bNominaEditarIngresosCargos = value;
			}
		}

		public bool BNominaEliminarExtras
		{
			get
			{
				return bNominaEliminarExtras;
			}
			set
			{
				bNominaEliminarExtras = value;
			}
		}

		public bool BNominaEliminarMovimientos
		{
			get
			{
				return bNominaEliminarMovimientos;
			}
			set
			{
				bNominaEliminarMovimientos = value;
			}
		}
		
		public bool BNominaEditarExtras
		{
			get
			{
				return bNominaEditarExtras;
			}
			set
			{
				bNominaEditarExtras = value;
			}
		}
		
		public bool BNominaEditarMovimientos
		{
			get
			{
				return bNominaEditarMovimientos;
			}
			set
			{
				bNominaEditarMovimientos = value;
			}
		}
		#endregion Nomina
		
		#region Cobranza
		public bool BReliquidar
		{
			get
			{
				return bReliquidar;
			}
			set
			{
				bReliquidar = value;
			}
		}

		public bool BReliquidarCruzandoCuentas
		{
			get
			{
				return bReliquidarCruzandoCuentas;
			}
			set
			{
				bReliquidarCruzandoCuentas = value;
			}
		}
		
		public bool B60DiasCobro
		{
			get
			{
				return b60DiasCobro;
			}
			set
			{
				b60DiasCobro = value;
			}
		}
		
		public bool BCruzarPagosAdelantados
		{
			get
			{
				return bCruzarPagosAdelantados;
			}
			set
			{
				bCruzarPagosAdelantados = value;
			}
		}

		public bool B30DiasCobro
		{
			get
			{
				return b30DiasCobro;
			}
			set
			{
				b30DiasCobro = value;
			}
		}
		
		public bool B15DiasCobro
		{
			get
			{
				return b15DiasCobro;
			}
			set
			{
				b15DiasCobro = value;
			}
		}
		
		public bool B1DiasCobro
		{
			get
			{
				return b1DiasCobro;
			}
			set
			{
				b1DiasCobro = value;
			}
		}
		public bool BDia
		{
			get
			{
				return bDia;
			}
			set
			{
				bDia = value;
			}
		}
		public bool BModificarCapitalInteres
		{
			get
			{
				return bModificarCapitalInteres;
			}
			set
			{
				bModificarCapitalInteres = value;
			}
		}
		#endregion Cobranza

		public bool BAsignaTecnicos 
		{
			get
			{
				return bAsignaTecnicos;
			}
			set
			{
				bAsignaTecnicos = value;
			}
		}
		
		public bool BModificaGarantia
		{
			get
			{
				return bModificaGarantia;
			}
			set
			{
				bModificaGarantia = value;
			}
		}

		public bool BVerRemates 
		{
			get
			{
				return bVerRemates;
			}
			set
			{
				bVerRemates = value;
			}
		}
		
		public bool BVerMatriz
		{
			get
			{
				return bVerMatriz;
			}
			set
			{
				bVerMatriz = value;
			}
		}

		public bool BModificarCuentas
		{
			get
			{
				return bModificarCuentas;
			}
			set
			{
				bModificarCuentas = value;
			}
		}

		public bool BEliminaModificaSeriales
		{
			get
			{
				return bEliminaModificaSeriales;
			}
			set
			{
				bEliminaModificaSeriales = value;
			}
		}

		public bool BCerrarCaja
		{
			get
			{
				return bCerrarCaja;
			}
			set
			{
				bCerrarCaja = value;
			}
		}

		public bool BAbrirCaja
		{
			get
			{
				return bAbrirCaja;
			}
			set
			{
				bAbrirCaja = value;
			}
		}

		public bool BLimiteBusquedaFechas
		{
			get
			{
				return bLimiteBusquedaFechas;
			}
			set
			{
				bLimiteBusquedaFechas = value;
			}
		}
		
		public bool BFormasDePagoCE
		{
			get
			{
				return bFormasDePagoCE;
			}
			set
			{
				bFormasDePagoCE = value;
			}
		}
		
		public bool BFormasDePagoCred
		{
			get
			{
				return bFormasDePagoCred;
			}
			set
			{
				bFormasDePagoCred = value;
			}
		}

		public bool BRevisadoCE
		{
			get
			{
				return bRevisadoCE;
			}
			set
			{
				bRevisadoCE = value;
			}
		}
 
		public bool BRefinanciar
		{
			get
			{
				return bRefinanciar;
			}
			set
			{
				bRefinanciar = value;
			}
		}

		public bool BCondonar
		{
			get
			{
				return bCondonar;
			}
			set
			{
				bCondonar = value;
			}
		}

		public bool BCobrar
		{
			get
			{
				return bCobrar;
			}
			set
			{
				bCobrar = value;
			}
		}

		public bool BHistorico
		{
			get
			{
				return bHistorico;
			}
			set
			{
				bHistorico = value;
			}
		}

		public bool BGenerarNC
		{
			get
			{
				return bGenerarNC;
			}
			set
			{
				bGenerarNC = value;
			}
		}

		public bool BGenerarNCP
		{
			get
			{
				return bGenerarNCP;
			}
			set
			{
				bGenerarNCP = value;
			}
		}

		public bool EditarLote
		{
			get
			{
				return editarLote;
			}
			set
			{
				editarLote = value;
			}
		}

		public bool EditarReferencia
		{
			get
			{
				return editarReferencia;
			}
			set
			{
				editarReferencia = value;
			}
		}

		public bool BArticulosAtributos
		{
			get
			{
				return bArticulosAtributos;
			}
			set
			{
				bArticulosAtributos = value;
			}
		}

		public bool BSolNCAFD
		{
			get
			{
				return bSolNCAFD;
			}
			set
			{
				bSolNCAFD = value;
			}
		}	

		public bool BFacturar
		{
			get
			{
				return bFacturar;
			}
			set
			{
				bFacturar = value;
			}
		}

		public bool BActualizar
		{
			get
			{
				return bActualizar;
			}
			set
			{
				bActualizar = value;
			}
		}

		public bool BGenerar
		{
			get
			{
				return bGenerar;
			}
			set
			{
				bGenerar = value;
			}
		}

		public bool BImportar
		{
			get
			{
				return bImportar;
			}
			set
			{
				bImportar = value;
			}
		}

		public bool BVerDocumentos
		{
			get
			{
				return bVerDocumentos;
			}
			set
			{
				bVerDocumentos = value;
			}
		}

		public bool BAyuda
		{
			get
			{
				return bAyuda;
			}
			set
			{
				bAyuda = value;
			}
		}

		public bool BAsignar
		{
			get
			{
				return bAsignar;
			}
			set
			{
				bAsignar = value;
			}
		}

		public bool BGuardar
		{
			get
			{
				return bGuardar;
			}
			set
			{
				bGuardar = value;
			}
		}

		public bool BVerMayor100dias
		{
			get
			{
				return bVerMayor100dias;
			}
			set
			{
				bVerMayor100dias = value;
			}
		}

		public bool BDescuento
		{
			get
			{
				return bDescuento;
			}
			set
			{
				bDescuento = value;
			}
		}

		public bool BAprobarValorMayor
		{
			get
			{
				return bAprobarValorMayor;
			}
			set
			{
				bAprobarValorMayor = value;
			}
		}

		public bool BSupervisor
		{
			get
			{
				return bSupervisor;
			}
			set
			{
				bSupervisor = value;
			}
		}

		public bool BAprobado
		{
			get
			{
				return bAprobado;
			}
			set
			{
				bAprobado = value;
			}
		}

		public bool BFecha3Días
		{
			get
			{
				return bFecha3Días;
			}
			set
			{
				bFecha3Días = value;
			}
		}

		public bool BEditPulgadas
		{
			get
			{
				return bEditPulgadas;
			}
			set
			{
				bEditPulgadas = value;
			}
		}

		public bool BEditColor
		{
			get
			{
				return bEditColor;
			}
			set
			{
				bEditColor = value;
			}
		}

		public bool BEditProcesador
		{
			get
			{
				return bEditProcesador;
			}
			set
			{
				bEditProcesador = value;
			}
		}

		public bool BEditMarca
		{
			get
			{
				return bEditMarca;
			}
			set
			{
				bEditMarca = value;
			}
		}

		public bool BEditMemoria
		{
			get
			{
				return bEditMemoria;
			}
			set
			{
				bEditMemoria = value;
			}
		}

		public bool BEditDisco
		{
			get
			{
				return bEditDisco;
			}
			set
			{
				bEditDisco = value;
			}
		}

		public bool BEditPrecio
		{
			get
			{
				return bEditPrecio;
			}
			set
			{
				bEditPrecio = value;
			}
		}

		public bool BEditGaming
		{
			get
			{
				return bEditGaming;
			}
			set
			{
				bEditGaming = value;
			}
		}

		public bool BEditLumenes
		{
			get
			{
				return bEditLumenes;
			}
			set
			{
				bEditLumenes = value;
			}
		}

		public bool BEditResolucion
		{
			get
			{
				return bEditResolucion;
			}
			set
			{
				bEditResolucion = value;
			}
		}

		public bool BEditCapacidad
		{
			get
			{
				return bEditCapacidad;
			}
			set
			{
				bEditCapacidad = value;
			}
		}

		public bool BEditAlmacenamiento
		{
			get
			{
				return bEditAlmacenamiento;
			}
			set
			{
				bEditAlmacenamiento = value;
			}
		}

		public bool BEditPixeles
		{
			get
			{
				return bEditPixeles;
			}
			set
			{
				bEditPixeles = value;
			}
		}

		public bool BEditQuemadores
		{
			get
			{
				return bEditQuemadores;
			}
			set
			{
				bEditQuemadores = value;
			}
		}

		public bool BEditTarjeta
		{
			get
			{
				return bEditTarjeta;
			}
			set
			{
				bEditTarjeta = value;
			}
		}

		public bool BEditOfeElectro
		{
			get
			{
				return bEditOfeElectro;
			}
			set
			{
				bEditOfeElectro = value;
			}
		}

		public bool BEditOfeComp
		{
			get
			{
				return bEditOfeComp;
			}
			set
			{
				bEditOfeComp = value;
			}
		}

		public bool BAccesoBoton
		{
			get
			{
				return bAccesoBoton;
			}
			set
			{
				bAccesoBoton = value;
			}
		}

		public bool BGestion
		{
			get
			{
				return bGestion;
			}
			set
			{
				bGestion = value;
			}
		}
		
		public bool BMora
		{
			get
			{
				return bMora;
			}
			set
			{
				bMora = value;
			}
		}
		
		public bool BFiniquitos
		{
			get
			{
				return bFiniquitos;
			}
			set
			{
				bFiniquitos = value;
			}
		}

		public bool BtnDescuento
		{
			get
			{
				return btnDescuento;
			}
			set
			{
				btnDescuento = value;
			}
		}

		public bool BtnDistribucion
		{
			get
			{
				return btnDistribucion;
			}
			set
			{
				btnDistribucion = value;
			}
		}
		public bool BDesModifica
		{
			get
			{
				return bDesModifica;
			}
			set
			{
				bDesModifica = value;
			}
		}

		public bool BDesEmpleado
		{
			get
			{
				return bDesEmpleado;
			}
			set
			{
				bDesEmpleado = value;
			}
		}

		public bool BRefacturacion
		{
			get
			{
				return bRefacturacion;
			}
			set
			{
				bRefacturacion = value;
			}
		}

		public bool BAnulaTiempo
		{
			get
			{
				return bAnulaTiempo;
			}
			set
			{
				bAnulaTiempo = value;
			}
		}
		
		public bool BEditacionMinutos
		{
			get
			{
				return bEditacionMinutos;
			}
			set
			{
				bEditacionMinutos = value;
			}
		}

		public bool BCobranzasSinLimiteFechas
		{
			get
			{
				return bCobranzasSinLimiteFechas;
			}
			set
			{
				bCobranzasSinLimiteFechas = value;
			}
		}

		public bool BTipoBono
		{
			get
			{
				return bTipoBono;
			}
			set
			{
				bTipoBono = value;
			}
		}

		public bool BAcreditar
		{
			get
			{
				return bAcreditar;
			}
			set
			{
				bAcreditar = value;
			}
		}

		public bool BEditaCantidad
		{
			get
			{
				return bEditaCantidad;
			}
			set
			{
				bEditaCantidad = value;
			}
		}
		
		public bool BSubirArchivos
		{
			get
			{
				return bSubirArchivos;
			}
			set
			{
				bSubirArchivos = value;
			}
		}
		
		public bool BImagen
		{
			get
			{
				return bImagen;
			}
			set
			{
				bImagen = value;
			}
		}

		public bool BCupon
		{
			get
			{
				return bCupon;
			}
			set
			{
				bCupon = value;
			}
		}
		public bool BHabilitaControles
		{
			get
			{
				return bHabilitaControles;
			}
			set
			{
				bHabilitaControles = value;
			}
		}

		public bool BFacturaExtras
		{
			get
			{
				return bFacturaExtras;
			}
			set
			{
				bFacturaExtras = value;
			}
		}

		public bool BFTemporal
		{
			get
			{
				return bFTemporal;
			}
			set
			{
				bFTemporal = value;
			}
		}

		public bool BCambiarPrecioOE
		{
			get
			{
				return bCambiarPrecioOE;
			}
			set
			{
				bCambiarPrecioOE = value;
			}
		}

		public bool BAutPagosValor
		{
			get
			{
				return bAutPagosValor;
			}
			set
			{
				bAutPagosValor = value;
			}
		}

		public bool BCuponEmpresas
		{
			get
			{
				return bCuponEmpresas;
			}
			set
			{
				bCuponEmpresas = value;
			}
		}

		public bool BCruzaAnticiposCC
		{
			get
			{
				return bCruzaAnticiposCC;
			}
			set
			{
				bCruzaAnticiposCC = value;
			}
		}
		
		public bool BAnulaCrucesProveedores
		{
			get
			{
				return bAnulaCrucesProveedores;
			}
			set
			{
				bAnulaCrucesProveedores = value;
			}
		}

		public bool BActivaTarjetaVIP
		{
			get
			{
				return bActivaTarjetaVIP;
			}
			set
			{
				bActivaTarjetaVIP = value;
			}
		}
		public bool ModificarMoraCartera
		{
			get
			{
				return modificarMoraCartera;
			}
			set 
			{
				modificarMoraCartera = value;
			}
		}
		public bool ModificarfechasAnteriores
		{
			get
			{
				return modificarfechasanteriores;
			}
			set 
			{
				modificarfechasanteriores = value;
			}
		}
		public bool BBuscarLista
		{
			get
			{
				return bBuscarLista;
			}
			set
			{
				bBuscarLista = value;
			}
		}
		public bool BExportarLista
		{
			get
			{
				return bExportaLista;
			}
			set
			{
				bExportaLista = value;
			}
		}
		public bool BHabilitaOPT
		{
			get
			{
				return bHabilitaOPT;
			}
			set
			{
				bHabilitaOPT = value;
			}
		}
		public bool AbreFacturacion
		{
			get
			{
				return abrefacturacion;
			}
			set 
			{
				abrefacturacion = value;
			}
		}
		public bool Responsable
		{
			get
			{
				return responsable;
			}
			set 
			{
				responsable = value;
			}
		}
		public bool EditarResponsable
		{
			get
			{
				return editarresponsable;
			}
			set 
			{
				editarresponsable = value;
			}
		}
		public bool CargoPersonal
		{
			get
			{
				return cargopersonal;
			}
			set 
			{
				cargopersonal = value;
			}
		}
		public bool AnulaCancelado
		{
			get
			{
				return bAnulaCancelado;
			}
			set 
			{
				bAnulaCancelado = value;
			}
		}
		public bool BAnulFacEnviada
		{
			get
			{
				return bAnulFacEnviada;
			}
			set 
			{
				bAnulFacEnviada = value;
			}
		}

		
		public Acceso (C1.Data.C1DataSet ds, string stCodigo)
		{
			string nivel = "n_" + MenuLatinium.iNivel.ToString().Trim();
			string stSelect = string.Format("Select IsNull(Pos, 0), IsNull({0}, 0) From Seguridad Where Codigo Like '{1}%'", nivel, stCodigo);
			SqlDataReader dr = Funcion.rEscalarSQL(ds, stSelect, true);
			while(dr.Read())
			{
				int iPos = dr.GetInt32(0);
				bool bHabilita = dr.GetBoolean(1);
				Habilita(iPos, bHabilita);
			}
			dr.Close();
		}

		private void Habilita(int iPos, bool bHabilita)
		{
			switch(iPos)
			{
				case 0:
					buscar = bHabilita;
					break;
				case 1:
					imprimir = bHabilita;
					break;
				case 2:
					nuevo = bHabilita;
					break;
				case 3:
					editar = bHabilita;
					break;
				case 4:
					editarSoloDia = bHabilita;
					break;
				case 5:
					editarDiaIngreso = bHabilita;
					break;
				case 6:
					anular = bHabilita;
					break;
				case 7:
					recuperar = bHabilita;
					break;
				case 8:
					eliminar = bHabilita;
					break;
				case 9:
					cambiarFecha = bHabilita;
					break;
				case 10:
					cambiarPeriodo = bHabilita;
					break;
				case 11:
					cambiarNumeracion = bHabilita;
					break;
				case 12:
					mostrarDia = bHabilita;
					break;
				case 14:
					imprimirRetencion = bHabilita;
					break;
				case 15:
					imprimirCheque = bHabilita;
					break;
				case 16:
					imprimirCheque = bHabilita;
					break;
				case 17:
					editarCuentaNombre = bHabilita;
					break;
				case 18:
					editarCuentaCodigo = bHabilita;
					break;
				case 19:
					editarTabla = bHabilita;
					break;
				case 20:
					crearAsiento = bHabilita;
					break;
				case 21:
					cambiarGrupo = bHabilita; // Opcion de seguridad
					break;
				case 22:
					transformacion = bHabilita; // Opcion de seguridad
					break;
				case 23:
					retencion = bHabilita; // Opcion de seguridad
					break;
				case 24:
					actualizaPrecio = bHabilita; // Opcion de seguridad
					break;
				case 25:
					cambiarPrecio = bHabilita;
					break;
				case 26:
					cambiarDescuento = bHabilita;
					break;
				case 27:
					ReImprimir = bHabilita;
					break;
				case 28:
					editaImpresa = bHabilita;
					break;
				case 29:
					grabaCero = bHabilita;
					break;
				case 30:
					autorizacion = bHabilita;
					break;
				case 31:
					listaPrecios = bHabilita;
					break;
				case 33:
					auditoria = bHabilita;
					break;
				case 34:
					EditarRevisadas = bHabilita;
					break;
				case 40:
					bEditarNombre = bHabilita;
					break;					
				case 831:
					EditarDetalleFormaPago = bHabilita;
					break;
				case 832:
					EditarContabilizado = bHabilita;
					break;
				case 838:
					BModificarGastos = bHabilita;
					break;
				case 839:
					BModificarMora = bHabilita;
					break;
				case 861:
					BPrecioDistribuidor = bHabilita;
					break;
				case 862:
					BAprobarPrefactura = bHabilita;
					break;
				case 863:
					BRechazarPrefactura = bHabilita;
					break;
				case 864:
					BBuscarOrdenCompra = bHabilita;
					break;	
				case 865:
					BFichaArticulos = bHabilita;
					break;
				case 866:
					bAuditoria = bHabilita;
					break;
				case 846:
					bDepositoCaja = bHabilita;
					break;
				case 847:
					bAutorizar = bHabilita;
					break;
				case 422:
					bAutorizarPedido = bHabilita;
					break;
				case 423:
					bVistas = bHabilita;
					break;
				case 425:
					bExportarExcel = bHabilita;
					break;
				case 426:
					bCambiarVencimiento = bHabilita;
					break;
				case 424:
					bModificarBodega = bHabilita;
					break;
				case 427:
					bAutEntrada = bHabilita;
					break;
				case 867:
					bAutCobro = bHabilita;
					break;
				case 37:
					bBuscarDocumentos = bHabilita;
					break;
				case 430:
					bActNum = bHabilita;
					break;
				case 431:
					bEditaCobranza = bHabilita;
					break;
				case 432:
					bAnulaCobranza = bHabilita;
					break;
				case 434:
					bModifcarFormaDePago = bHabilita;
					break;
				case 435:
					bContabilizarFactMan = bHabilita;
					break;
				case 444:
					bVerProveedorDT = bHabilita;
					break;
				case 460:
					bSobrantesFaltantesCaja = bHabilita;
					break;
				case 462:
					bModificaRenta = bHabilita;
					break;
				case 463:
					bModificaIVA = bHabilita;
					break;
					#region Autorizacion de Pagos
				case 453:
					bAPAutorizar = bHabilita;
					break;
				case 454:
					bAPAsignaBancos = bHabilita;
					break;
				case 455:
					bAPActAutorizados = bHabilita;
					break;
				case 456:
					bAPVerAutorizados = bHabilita;
					break;
				case 457:
					bAPVerPagados = bHabilita;
					break;
					#endregion Autorizacion de Pagos
				case 471:
					bModArtCmbProm = bHabilita;
					break;
				case 473:
					bEliArtCmbProm = bHabilita;
					break;
				case 475:
					bEliArt = bHabilita;
					break;
				case 477:
					bBloquearCB = bHabilita;
					break;
				case 478:
					bDesbloquearCB = bHabilita;
					break;
				case 483:
					bEditarProcesado = bHabilita;
					break;
				case 484:
					bModArtSerializado = bHabilita;
					break;
				case 485:
					bEliArtSerializado = bHabilita;
					break;
				case 487:
					BAnularProcesado = bHabilita;
					break;
				case 488:
					bAnularConSeriales = bHabilita;
					break;
				case 492:
					bEditarAprobado = bHabilita;
					break;
				case 490:
					bAnularAprobado = bHabilita;
					break;
				case 519:
					bEditarPagosAutorizados = bHabilita;
					break;
				case 522:
					bEditarPagosRealizados = bHabilita;
					break;
				case 520:
					bEditarLiquidado = bHabilita;
					break;
				case 521:
					bAnularPagosAutorizados = bHabilita;
					break;
				case 523:
					bAnularPagosRealizados = bHabilita;
					break;
				case 508:
					bAnularLiquidado = bHabilita;
					break;
				case 524:
					bEliminaFormaDePago = bHabilita;
					break;
				case 546:
					bManejaSerial = bHabilita;
					break;
				case 547:
					bCruzaAnticiposCP = bHabilita;
					break;	
				case 548:
					bEliminaSeriales = bHabilita;
					break;	
				case 549:
					bModificaSeriales = bHabilita;
					break;	
				case 550:
					bVerSerialesProcesadas = bHabilita;
					break;	
				case 565:
					bProcesarSerialesFaltantes = bHabilita;
					break;	
				case 566:
					bModificarSaldoCB = bHabilita;
					break;
				case 567:
					bAsignaCobrador = bHabilita;
					break;
				case 364:
					bAnulaSoloDelDia = bHabilita;
					break;
				case 572:
					bEliminaFilaCCEnEspera = bHabilita;
					break;
				case 573:
					bEliminaFilaCCProcesada = bHabilita;
					break;
				case 576:
					bProcesar = bHabilita;
					break;
				case 583:
					bAnularEnTransito = bHabilita;
					break;
				case 584:
					bAnularTFProcesado = bHabilita;
					break;
				case 588:
					bModificarImpreso = bHabilita;
					break;
				case 590:
					bAbrirImpuestos = bHabilita;
					break;
				case 591:
					bVerCapitalInteres = bHabilita;
					break;
				case 636:
					bCarpa = bHabilita;
					break;
				case 638:
					bExportar = bHabilita;
					break;					
				case 642:
					bTodos = bHabilita;
					break;
				case 643:
					bDepartamentoTecnico = bHabilita;
					break;
				case 644:
					bCompras = bHabilita;
					break;
				case 645:
					bEstados = bHabilita;
					break;
				case 647:
					bTransferir = bHabilita;
					break;
				case 648:
					bModRevDT = bHabilita;
					break;
				case 754:
					bModRevCompras = bHabilita;
					break;
				case 755:
					bCotizaAccesorios = bHabilita;
					break;
				case 756:
					bRevisionObsoleto = bHabilita;
					break;					
				case 660:					
					bCom_FiltrarPorCargo = bHabilita;
					break;
				case 661:
					bCom_CerrarPeriodo = bHabilita;
					break;
				case 664:
					bCom_Observaciones = bHabilita;
					break;
				case 677:
					bVerComisiones = bHabilita;
					break;
					#region Rol de Pagos
				case 268:
					bRolContabilizar = bHabilita;
					break;
				case 665:
					bRolGenerar = bHabilita;
					break;
				case 673:
					bRolRevision = bHabilita;
					break;
				case 667:
					bRolQuitarRevisado = bHabilita;
					break;
				case 668:
					bRolAutorizar = bHabilita;
					break;
				case 669:
					bRolQuitarAutorizado = bHabilita;
					break;
				case 670:
					bRolAcreditar = bHabilita;
					break;						
				case 674:
					bRolSubirArchivos = bHabilita;
					break;	
					#endregion Rol de Pagos
				case 681:
					bGenerarCobro = bHabilita;
					break;
					#region Nomina
				case 691:
					bNominaEditarIngresosCargos = bHabilita;
					break;
				case 692:
					bNominaEliminarExtras = bHabilita;
					break;
				case 693:
					bNominaEliminarMovimientos = bHabilita;
					break;
				case 694:
					bNominaEditarExtras = bHabilita;
					break;
				case 695:
					bNominaEditarMovimientos = bHabilita;
					break;
					#endregion Nomina
					#region Cobranzas
				case 700:
					bReliquidar = bHabilita;
					break;
				case 701:
					bReliquidarCruzandoCuentas = bHabilita;
					break;					
				case 783:
					b60DiasCobro = bHabilita;
					break;
				case 784:
					bCruzarPagosAdelantados = bHabilita;
					break;
				case 702:
					b30DiasCobro = bHabilita;
					break;
				case 703:
					b15DiasCobro = bHabilita;
					break;
				case 704:
					bModificarCapitalInteres = bHabilita;
					break;					
					#endregion Cobranzas
				case 716:
					bAprobarCarpa = bHabilita;
					break;
				case 721:
					bIrA = bHabilita;
					break;
				case 762:
					bAsignaTecnicos = bHabilita;
					break;
				case 765:
					bModificaGarantia = bHabilita;
					break;
				case 768:
					bVerRemates = bHabilita;
					break;
				case 769:
					bVerMatriz = bHabilita;
					break;
				case 773:
					bModificarCuentas = bHabilita;
					break;
				case 778:
					bEliminaModificaSeriales = bHabilita;
					break;
				case 785:
					bCerrarCaja = bHabilita;
					break;
				case 786:
					bAbrirCaja = bHabilita;
					break;
				case 811:
					anularCP = bHabilita;
					break;
				case 852:
					bLimiteBusquedaFechas = bHabilita;
					break;
				case 882:
					bFormasDePagoCE = bHabilita;
					break;
				case 437:
					bFormasDePagoCred = bHabilita;
					break;	
				case 893:
					bRevisadoCE = bHabilita;
					break;
				case 965:
					bRefinanciar = bHabilita;
					break;
				case 966:
					bCondonar = bHabilita;
					break;
				case 967:
					bCobrar = bHabilita;
					break;	
				case 968:
					bHistorico = bHabilita;
					break;
				case 1209:
					bGenerarNC = bHabilita;
					break;	
				case 1210:
					bGenerarNCP = bHabilita;
					break;
				case 1211:
					b1DiasCobro = bHabilita;
					break;
				case 1212:
					bDia = bHabilita; 
					break;
				case 1213:
					CambiarFechaPosterior = bHabilita;
					break;
				case 1214:
					editarLote = bHabilita;
					break;
				case 1215:
					editarReferencia = bHabilita;
					break;
				case 1216:
					bArticulosAtributos = bHabilita;
					break;
				case 1217:
					bFacturar = bHabilita;
					break;
				case 1218:
					bActualizar = bHabilita;
					break;
				case 1219:
					bGenerar = bHabilita;
					break;
				case 1220:
					bImportar = bHabilita;
					break;
				case 1221:
					bVerDocumentos = bHabilita;
					break;
				case 1222:
					bAyuda = bHabilita;
					break;
				case 1487:
					bSolNCAFD = bHabilita;
					break;
				case 1223:
					bAsignar = bHabilita;
					break;
				case 1224:
					bGuardar = bHabilita;
					break;
				case 1225:
					bDescuento = bHabilita;
					break;
				case 1564:
					bAprobarValorMayor = bHabilita;
					break;
				case 1565:
					bVerMayor100dias = bHabilita;
					break;
				case 1566:
					bSupervisor = bHabilita;
					break;
				case 1567:
					bAprobado = bHabilita;
					break;
				case 1568:
					bFecha3Días = bHabilita;
					break;
				case 1569:
					bEditPulgadas = bHabilita;
					break;
				case 1570:
					bEditColor = bHabilita;
					break;
				case 1571:
					bEditProcesador = bHabilita;
					break;
				case 1572:
					bEditMarca = bHabilita;
					break;
				case 1573:
					bEditMemoria = bHabilita;
					break;
				case 1574:
					bEditDisco = bHabilita;
					break;
				case 1575:
					bEditPrecio = bHabilita;
					break;
				case 1576:
					bEditGaming = bHabilita;
					break;
				case 1577:
					bEditLumenes = bHabilita;
					break;
				case 1578:
					bEditResolucion = bHabilita;
					break;
				case 1579:
					bEditCapacidad = bHabilita;
					break;
				case 1580:
					bEditAlmacenamiento = bHabilita;
					break;
				case 1581:
					bEditPixeles = bHabilita;
					break;
				case 1582:
					bEditQuemadores = bHabilita;
					break;
				case 1583:
					bEditTarjeta = bHabilita;
					break;
				case 1584:
					bEditOfeElectro = bHabilita;
					break;
				case 1585:
					bEditOfeComp = bHabilita;
					break;
				case 1586:
					bAccesoBoton = bHabilita;
					break;
				case 1587:
					bAnulaFacturaMayor0 = bHabilita;
					break;
				case 1588:
					bGestion = bHabilita;
					break;
				case 1589:
					bMora = bHabilita;
					break;
				case 1710:
					bFiniquitos = bHabilita;
					break;
				case 1711:
					btnDescuento = bHabilita;
					break;
				case 1712:
					btnDistribucion = bHabilita;
					break;
				case 1713:
					bDesModifica = bHabilita;
					break;
				case 1714:
					bDesEmpleado = bHabilita;
					break;
				case 1715:
					bRefacturacion = bHabilita;
					break;
				case 1716:
					bAnulaTiempo = bHabilita;
					break;
				case 1717:
					bEditacionMinutos = bHabilita;
					break;	
				case 1742:
					bCobranzasSinLimiteFechas = bHabilita;
					break;
				case 1743:
					bTipoBono = bHabilita;
					break;
				case 1744:
					bAcreditar = bHabilita;
					break;
				case 1745:
					bEditaCantidad = bHabilita;
					break;
				case 1746:
					bSubirArchivos = bHabilita;
					break;					
				case 1747:
					bImagen = bHabilita;
					break;
				case 1748:
					bCupon = bHabilita;
					break;
				case 1749:
					bHabilitaControles = bHabilita;
					break;
				case 1750:
					bFacturaExtras = bHabilita;
					break;
				case 1751:
					bFTemporal = bHabilita;
					break;
				case 1752:
					bCambiarPrecioOE = bHabilita;
					break;
				case 1753:
					bAutPagosValor = bHabilita;
					break;
				case 1754:
					bCuponEmpresas = bHabilita;
					break;
				case 1755:
					bCruzaAnticiposCC = bHabilita;
					break;
				case 1991:
					bAnulaCrucesProveedores = bHabilita;
					break;
				case 1992:
					bActivaTarjetaVIP = bHabilita;
					break;
				case 1993:
					anularDocSRI = bHabilita;
					break;
				case 1994:
					cambiarFechaDiciembre = bHabilita;
					break;
				case 2142:
					modificarMoraCartera = bHabilita;
					break;
				case 2143:
					modificarfechasanteriores = bHabilita;
					break;
				case 2146:					
					bCom_FiltrarPorCargo = bHabilita;
					break;
				case 2147:
					bCom_CerrarPeriodo = bHabilita;
					break;
				case 2148:
					bCom_Observaciones = bHabilita;
					break;
				case 2149:
					bVerComisiones = bHabilita;
					break;
				case 2153:
					bBuscarLista = bHabilita;
					break;
				case 2154:
					bHabilitaOPT = bHabilita;
					break;
				case 2155:
					bExportaLista = bHabilita;
					break;
				case 2156:
					AbreFacturacion= bHabilita;
					break;
				case 2157:
					Responsable= bHabilita;
					break;
				case 2159:
					EditarResponsable= bHabilita;
					break;
				case 2176:
					CargoPersonal = bHabilita;
					break;
				case 2191:
					BExportar = bHabilita;
					break;
				case 2192:
					BVerDocumentos = bHabilita;
					break;
				case 2193:
					AnulaCancelado = bHabilita;
					break;
				case 2194:
					bAnulFacEnviada = bHabilita;
					break;
			}
		}		

		/*  
		 *  f
		private bool bGenerarNC = true;
		private bool  = true;

		 * private bool bRefinanciar = true;
		private bool bCondonar = true;
		private bool bCobrar = true;
		private bool bHistorico = true;
		965	REFINANCIAR
966	CONDONAR
967	COBRAR
968	HISTORICO*/
		public Acceso (string stCodigo, DataTable dtPermiso)
		{
			DataRow [] drCod;
			drCod = dtPermiso.Select("codigo like '" + stCodigo + "*'", "pos");
			string nivel = "n_" + MenuLatinium.iNivel.ToString().Trim();
			foreach(DataRow dr in drCod)
			{
				if (dr["pos"] == System.DBNull.Value) continue;
				bool bHabilita = bool.Parse(dr[nivel].ToString());
				Habilita(int.Parse(dr["pos"].ToString()), bHabilita);
			}
		}
	}
}
