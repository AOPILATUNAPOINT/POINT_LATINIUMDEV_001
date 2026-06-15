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
	/// Descripción breve de frmProyeccionPresupuestoAlmacen.
	/// </summary>
	public class frmProyeccionPresupuestoAlmacen : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label40;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneral;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneralEnero;
		private System.Windows.Forms.Label label39;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFDGEnero;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrestamosEnero;
		private System.Windows.Forms.Label label37;
		private System.Windows.Forms.Label label36;
		private System.Windows.Forms.Label label35;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.Label label34;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGNCEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMNCEnero;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label label32;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSMAGastosRetirosDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSMAGastosRetirosNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSMAGastosRetirosOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSMAGastosRetirosSeptiembre;
		private System.Windows.Forms.Button btnActualizar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSMAGastosRetirosAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSMAGastosRetirosJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSMAGastosRetirosJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSMAGastosRetirosMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSMAGastosRetirosAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSMAGastosRetirosMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSMAGastosRetirosFebrero;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label31;
		private System.Windows.Forms.Label label30;
		private System.Windows.Forms.Label label29;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastosEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtResultadoEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosRetiroEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresosEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobranzaEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTotalEnero;
		private System.Windows.Forms.Label label25;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTotalEnero;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenCreditoFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenTarjetaFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenEfectivoFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMargenDistribuidorFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaCreditoFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaTarjetaFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaEfectivoFebrero;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionDiciembreVenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionNoviembreVenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionOctubreVenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionSeptiembreVenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionAgostoVenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionJulioVenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionJunioVenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMayoVenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionAbrilVenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionMarzoVenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionVentaDistribuidorFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionFebreroVenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyeccionEneroVenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.ToolTip toolTip1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.ComponentModel.IContainer components;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAño;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label27;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
    
		int iPeriodo = 0;
		private System.Windows.Forms.Label label41;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optContrato;
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		public frmProyeccionPresupuestoAlmacen()
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
		DateTime dtFechaInicial = DateTime.Today;
		DateTime dtFechaFinal = DateTime.Today;
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEnero", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnFebrero", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMarzo", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAbril", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMayo", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJunio", 5);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJulio", 6);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAgosto", 7);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeptiembre", 8);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnOctubre", 9);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnNoviembre", 10);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDiciembre", 11);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 3, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 4, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 5, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 6, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 7, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 8, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 9, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 10, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 11, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 12, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 13, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 14, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 15, true);
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmProyeccionPresupuestoAlmacen));
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance155 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance156 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance157 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance158 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance159 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance160 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance161 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance162 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance163 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance164 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance165 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance166 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance167 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance168 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance169 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance170 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance171 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance172 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance173 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance174 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance175 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance176 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance177 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance178 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance179 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance180 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance181 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance182 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance183 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance184 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance185 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance186 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance187 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance188 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance189 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance190 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance191 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance192 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance193 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance194 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance195 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance196 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance197 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance198 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance199 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance200 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance201 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance202 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance203 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance204 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance205 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance206 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance207 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance208 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance209 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance210 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance211 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance212 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance213 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance214 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance215 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance216 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance217 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance218 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance219 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance220 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance221 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance222 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance223 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance224 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance225 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance226 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance227 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance228 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance229 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance230 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance231 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance232 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance233 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance234 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance235 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance236 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance237 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance238 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance239 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance240 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance241 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance242 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance243 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance244 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance245 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance246 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance247 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance248 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance249 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance250 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance251 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance252 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance253 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance254 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance255 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance256 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance257 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance258 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance259 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance260 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance261 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance262 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance263 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance264 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance265 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance266 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance267 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance268 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance269 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance270 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance271 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance272 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance273 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance274 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance275 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance276 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance277 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance278 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance279 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance280 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance281 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance282 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance283 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance284 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance285 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance286 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance287 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance288 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance289 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance290 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance291 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance292 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance293 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance294 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance295 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance296 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance297 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance298 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance299 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance300 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance301 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance302 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance303 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance304 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance305 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance306 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance307 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.label40 = new System.Windows.Forms.Label();
			this.txtTotalGeneral = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneralEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label39 = new System.Windows.Forms.Label();
			this.txtFDGTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFDGEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label38 = new System.Windows.Forms.Label();
			this.txtPrestamosTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrestamosEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label37 = new System.Windows.Forms.Label();
			this.label36 = new System.Windows.Forms.Label();
			this.label35 = new System.Windows.Forms.Label();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label33 = new System.Windows.Forms.Label();
			this.label34 = new System.Windows.Forms.Label();
			this.txtGNCTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGNCEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMNCEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnExcel = new System.Windows.Forms.Button();
			this.label32 = new System.Windows.Forms.Label();
			this.txtSMAGastosRetirosDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSMAGastosRetirosNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSMAGastosRetirosOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSMAGastosRetirosSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.txtSMAGastosRetirosAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSMAGastosRetirosJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSMAGastosRetirosJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSMAGastosRetirosMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSMAGastosRetirosAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSMAGastosRetirosMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSMAGastosRetirosFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.label31 = new System.Windows.Forms.Label();
			this.label30 = new System.Windows.Forms.Label();
			this.label29 = new System.Windows.Forms.Label();
			this.label28 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtTotalGastos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastosEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtResultadoEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosRetiroEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalIngresosEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobranzaEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTotalEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label25 = new System.Windows.Forms.Label();
			this.txtProyeccionVentaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTotalEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtProyeccionMargenCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenCreditoFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenTarjetaFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenEfectivoFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMargenDistribuidorFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaCreditoFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaTarjetaFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaEfectivoFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtProyeccionVentaDistribuidor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionDiciembreVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionNoviembreVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionOctubreVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionSeptiembreVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionAgostoVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionJulioVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionJunioVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMayoVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionAbrilVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionMarzoVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionVentaDistribuidorFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionFebreroVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProyeccionEneroVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtAño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.label27 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label41 = new System.Windows.Forms.Label();
			this.optContrato = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneral)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionDiciembreVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionNoviembreVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionOctubreVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionSeptiembreVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionAgostoVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionJulioVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionJunioVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMayoVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionAbrilVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMarzoVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionFebreroVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionEneroVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optContrato)).BeginInit();
			this.SuspendLayout();
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label40.ForeColor = System.Drawing.Color.Black;
			this.label40.Location = new System.Drawing.Point(16, 712);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(22, 16);
			this.label40.TabIndex = 1455;
			this.label40.Text = "(=) ";
			this.label40.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotalGeneral
			// 
			appearance1.BackColorDisabled = System.Drawing.Color.White;
			appearance1.BackColorDisabled2 = System.Drawing.Color.White;
			appearance1.BorderColor = System.Drawing.Color.Black;
			appearance1.BorderColor3DBase = System.Drawing.Color.Black;
			appearance1.FontData.BoldAsString = "True";
			appearance1.ForeColor = System.Drawing.Color.Firebrick;
			appearance1.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneral.Appearance = appearance1;
			this.txtTotalGeneral.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneral.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneral.Enabled = false;
			this.txtTotalGeneral.FormatString = "#,##0.00";
			this.txtTotalGeneral.Location = new System.Drawing.Point(1114, 712);
			this.txtTotalGeneral.Name = "txtTotalGeneral";
			this.txtTotalGeneral.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneral.PromptChar = ' ';
			this.txtTotalGeneral.Size = new System.Drawing.Size(100, 19);
			this.txtTotalGeneral.TabIndex = 1454;
			// 
			// txtTotalGeneralNoviembre
			// 
			appearance2.BackColorDisabled = System.Drawing.Color.White;
			appearance2.BackColorDisabled2 = System.Drawing.Color.White;
			appearance2.BorderColor = System.Drawing.Color.Black;
			appearance2.BorderColor3DBase = System.Drawing.Color.Black;
			appearance2.ForeColor = System.Drawing.Color.Firebrick;
			appearance2.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralNoviembre.Appearance = appearance2;
			this.txtTotalGeneralNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralNoviembre.Enabled = false;
			this.txtTotalGeneralNoviembre.FormatString = "#,##0.00";
			this.txtTotalGeneralNoviembre.Location = new System.Drawing.Point(938, 712);
			this.txtTotalGeneralNoviembre.Name = "txtTotalGeneralNoviembre";
			this.txtTotalGeneralNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralNoviembre.PromptChar = ' ';
			this.txtTotalGeneralNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralNoviembre.TabIndex = 1453;
			// 
			// txtTotalGeneralDiciembre
			// 
			appearance3.BackColorDisabled = System.Drawing.Color.White;
			appearance3.BackColorDisabled2 = System.Drawing.Color.White;
			appearance3.BorderColor = System.Drawing.Color.Black;
			appearance3.BorderColor3DBase = System.Drawing.Color.Black;
			appearance3.ForeColor = System.Drawing.Color.Firebrick;
			appearance3.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralDiciembre.Appearance = appearance3;
			this.txtTotalGeneralDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralDiciembre.Enabled = false;
			this.txtTotalGeneralDiciembre.FormatString = "#,##0.00";
			this.txtTotalGeneralDiciembre.Location = new System.Drawing.Point(1026, 712);
			this.txtTotalGeneralDiciembre.Name = "txtTotalGeneralDiciembre";
			this.txtTotalGeneralDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralDiciembre.PromptChar = ' ';
			this.txtTotalGeneralDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralDiciembre.TabIndex = 1452;
			// 
			// txtTotalGeneralSeptiembre
			// 
			appearance4.BackColorDisabled = System.Drawing.Color.White;
			appearance4.BackColorDisabled2 = System.Drawing.Color.White;
			appearance4.BorderColor = System.Drawing.Color.Black;
			appearance4.BorderColor3DBase = System.Drawing.Color.Black;
			appearance4.ForeColor = System.Drawing.Color.Firebrick;
			appearance4.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralSeptiembre.Appearance = appearance4;
			this.txtTotalGeneralSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralSeptiembre.Enabled = false;
			this.txtTotalGeneralSeptiembre.FormatString = "#,##0.00";
			this.txtTotalGeneralSeptiembre.Location = new System.Drawing.Point(762, 712);
			this.txtTotalGeneralSeptiembre.Name = "txtTotalGeneralSeptiembre";
			this.txtTotalGeneralSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralSeptiembre.PromptChar = ' ';
			this.txtTotalGeneralSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralSeptiembre.TabIndex = 1451;
			// 
			// txtTotalGeneralOctubre
			// 
			appearance5.BackColorDisabled = System.Drawing.Color.White;
			appearance5.BackColorDisabled2 = System.Drawing.Color.White;
			appearance5.BorderColor = System.Drawing.Color.Black;
			appearance5.BorderColor3DBase = System.Drawing.Color.Black;
			appearance5.ForeColor = System.Drawing.Color.Firebrick;
			appearance5.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralOctubre.Appearance = appearance5;
			this.txtTotalGeneralOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralOctubre.Enabled = false;
			this.txtTotalGeneralOctubre.FormatString = "#,##0.00";
			this.txtTotalGeneralOctubre.Location = new System.Drawing.Point(850, 712);
			this.txtTotalGeneralOctubre.Name = "txtTotalGeneralOctubre";
			this.txtTotalGeneralOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralOctubre.PromptChar = ' ';
			this.txtTotalGeneralOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralOctubre.TabIndex = 1450;
			// 
			// txtTotalGeneralJulio
			// 
			appearance6.BackColorDisabled = System.Drawing.Color.White;
			appearance6.BackColorDisabled2 = System.Drawing.Color.White;
			appearance6.BorderColor = System.Drawing.Color.Black;
			appearance6.BorderColor3DBase = System.Drawing.Color.Black;
			appearance6.ForeColor = System.Drawing.Color.Firebrick;
			appearance6.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralJulio.Appearance = appearance6;
			this.txtTotalGeneralJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralJulio.Enabled = false;
			this.txtTotalGeneralJulio.FormatString = "#,##0.00";
			this.txtTotalGeneralJulio.Location = new System.Drawing.Point(586, 712);
			this.txtTotalGeneralJulio.Name = "txtTotalGeneralJulio";
			this.txtTotalGeneralJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralJulio.PromptChar = ' ';
			this.txtTotalGeneralJulio.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralJulio.TabIndex = 1449;
			// 
			// txtTotalGeneralAgosto
			// 
			appearance7.BackColorDisabled = System.Drawing.Color.White;
			appearance7.BackColorDisabled2 = System.Drawing.Color.White;
			appearance7.BorderColor = System.Drawing.Color.Black;
			appearance7.BorderColor3DBase = System.Drawing.Color.Black;
			appearance7.ForeColor = System.Drawing.Color.Firebrick;
			appearance7.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralAgosto.Appearance = appearance7;
			this.txtTotalGeneralAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralAgosto.Enabled = false;
			this.txtTotalGeneralAgosto.FormatString = "#,##0.00";
			this.txtTotalGeneralAgosto.Location = new System.Drawing.Point(674, 712);
			this.txtTotalGeneralAgosto.Name = "txtTotalGeneralAgosto";
			this.txtTotalGeneralAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralAgosto.PromptChar = ' ';
			this.txtTotalGeneralAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralAgosto.TabIndex = 1448;
			// 
			// txtTotalGeneralMayo
			// 
			appearance8.BackColorDisabled = System.Drawing.Color.White;
			appearance8.BackColorDisabled2 = System.Drawing.Color.White;
			appearance8.BorderColor = System.Drawing.Color.Black;
			appearance8.BorderColor3DBase = System.Drawing.Color.Black;
			appearance8.ForeColor = System.Drawing.Color.Firebrick;
			appearance8.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralMayo.Appearance = appearance8;
			this.txtTotalGeneralMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralMayo.Enabled = false;
			this.txtTotalGeneralMayo.FormatString = "#,##0.00";
			this.txtTotalGeneralMayo.Location = new System.Drawing.Point(410, 712);
			this.txtTotalGeneralMayo.Name = "txtTotalGeneralMayo";
			this.txtTotalGeneralMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralMayo.PromptChar = ' ';
			this.txtTotalGeneralMayo.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralMayo.TabIndex = 1447;
			// 
			// txtTotalGeneralJunio
			// 
			appearance9.BackColorDisabled = System.Drawing.Color.White;
			appearance9.BackColorDisabled2 = System.Drawing.Color.White;
			appearance9.BorderColor = System.Drawing.Color.Black;
			appearance9.BorderColor3DBase = System.Drawing.Color.Black;
			appearance9.ForeColor = System.Drawing.Color.Firebrick;
			appearance9.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralJunio.Appearance = appearance9;
			this.txtTotalGeneralJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralJunio.Enabled = false;
			this.txtTotalGeneralJunio.FormatString = "#,##0.00";
			this.txtTotalGeneralJunio.Location = new System.Drawing.Point(498, 712);
			this.txtTotalGeneralJunio.Name = "txtTotalGeneralJunio";
			this.txtTotalGeneralJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralJunio.PromptChar = ' ';
			this.txtTotalGeneralJunio.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralJunio.TabIndex = 1446;
			// 
			// txtTotalGeneralMarzo
			// 
			appearance10.BackColorDisabled = System.Drawing.Color.White;
			appearance10.BackColorDisabled2 = System.Drawing.Color.White;
			appearance10.BorderColor = System.Drawing.Color.Black;
			appearance10.BorderColor3DBase = System.Drawing.Color.Black;
			appearance10.ForeColor = System.Drawing.Color.Firebrick;
			appearance10.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralMarzo.Appearance = appearance10;
			this.txtTotalGeneralMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralMarzo.Enabled = false;
			this.txtTotalGeneralMarzo.FormatString = "#,##0.00";
			this.txtTotalGeneralMarzo.Location = new System.Drawing.Point(234, 712);
			this.txtTotalGeneralMarzo.Name = "txtTotalGeneralMarzo";
			this.txtTotalGeneralMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralMarzo.PromptChar = ' ';
			this.txtTotalGeneralMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralMarzo.TabIndex = 1445;
			// 
			// txtTotalGeneralAbril
			// 
			appearance11.BackColorDisabled = System.Drawing.Color.White;
			appearance11.BackColorDisabled2 = System.Drawing.Color.White;
			appearance11.BorderColor = System.Drawing.Color.Black;
			appearance11.BorderColor3DBase = System.Drawing.Color.Black;
			appearance11.ForeColor = System.Drawing.Color.Firebrick;
			appearance11.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralAbril.Appearance = appearance11;
			this.txtTotalGeneralAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralAbril.Enabled = false;
			this.txtTotalGeneralAbril.FormatString = "#,##0.00";
			this.txtTotalGeneralAbril.Location = new System.Drawing.Point(322, 712);
			this.txtTotalGeneralAbril.Name = "txtTotalGeneralAbril";
			this.txtTotalGeneralAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralAbril.PromptChar = ' ';
			this.txtTotalGeneralAbril.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralAbril.TabIndex = 1444;
			// 
			// txtTotalGeneralFebrero
			// 
			appearance12.BackColorDisabled = System.Drawing.Color.White;
			appearance12.BackColorDisabled2 = System.Drawing.Color.White;
			appearance12.BorderColor = System.Drawing.Color.Black;
			appearance12.BorderColor3DBase = System.Drawing.Color.Black;
			appearance12.ForeColor = System.Drawing.Color.Firebrick;
			appearance12.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralFebrero.Appearance = appearance12;
			this.txtTotalGeneralFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralFebrero.Enabled = false;
			this.txtTotalGeneralFebrero.FormatString = "#,##0.00";
			this.txtTotalGeneralFebrero.Location = new System.Drawing.Point(146, 712);
			this.txtTotalGeneralFebrero.Name = "txtTotalGeneralFebrero";
			this.txtTotalGeneralFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralFebrero.PromptChar = ' ';
			this.txtTotalGeneralFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralFebrero.TabIndex = 1443;
			// 
			// txtTotalGeneralEnero
			// 
			appearance13.BackColorDisabled = System.Drawing.Color.White;
			appearance13.BackColorDisabled2 = System.Drawing.Color.White;
			appearance13.BorderColor = System.Drawing.Color.Black;
			appearance13.BorderColor3DBase = System.Drawing.Color.Black;
			appearance13.ForeColor = System.Drawing.Color.Firebrick;
			appearance13.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalGeneralEnero.Appearance = appearance13;
			this.txtTotalGeneralEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGeneralEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneralEnero.Enabled = false;
			this.txtTotalGeneralEnero.FormatString = "#,##0.00";
			this.txtTotalGeneralEnero.Location = new System.Drawing.Point(58, 712);
			this.txtTotalGeneralEnero.Name = "txtTotalGeneralEnero";
			this.txtTotalGeneralEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneralEnero.PromptChar = ' ';
			this.txtTotalGeneralEnero.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGeneralEnero.TabIndex = 1442;
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label39.ForeColor = System.Drawing.Color.Black;
			this.label39.Location = new System.Drawing.Point(8, 688);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(39, 16);
			this.label39.TabIndex = 1441;
			this.label39.Text = "+ FDG";
			this.label39.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFDGTotal
			// 
			appearance14.BackColorDisabled = System.Drawing.Color.White;
			appearance14.BackColorDisabled2 = System.Drawing.Color.White;
			appearance14.BorderColor = System.Drawing.Color.Black;
			appearance14.BorderColor3DBase = System.Drawing.Color.Black;
			appearance14.FontData.BoldAsString = "True";
			appearance14.ForeColor = System.Drawing.Color.Firebrick;
			appearance14.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGTotal.Appearance = appearance14;
			this.txtFDGTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGTotal.Enabled = false;
			this.txtFDGTotal.FormatString = "#,##0.00";
			this.txtFDGTotal.Location = new System.Drawing.Point(1114, 688);
			this.txtFDGTotal.Name = "txtFDGTotal";
			this.txtFDGTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGTotal.PromptChar = ' ';
			this.txtFDGTotal.Size = new System.Drawing.Size(100, 19);
			this.txtFDGTotal.TabIndex = 1440;
			// 
			// txtFDGNoviembre
			// 
			appearance15.BackColorDisabled = System.Drawing.Color.White;
			appearance15.BackColorDisabled2 = System.Drawing.Color.White;
			appearance15.BorderColor = System.Drawing.Color.Black;
			appearance15.BorderColor3DBase = System.Drawing.Color.Black;
			appearance15.ForeColor = System.Drawing.Color.Firebrick;
			appearance15.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGNoviembre.Appearance = appearance15;
			this.txtFDGNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGNoviembre.Enabled = false;
			this.txtFDGNoviembre.FormatString = "#,##0.00";
			this.txtFDGNoviembre.Location = new System.Drawing.Point(938, 688);
			this.txtFDGNoviembre.Name = "txtFDGNoviembre";
			this.txtFDGNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGNoviembre.PromptChar = ' ';
			this.txtFDGNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtFDGNoviembre.TabIndex = 1439;
			// 
			// txtFDGDiciembre
			// 
			appearance16.BackColorDisabled = System.Drawing.Color.White;
			appearance16.BackColorDisabled2 = System.Drawing.Color.White;
			appearance16.BorderColor = System.Drawing.Color.Black;
			appearance16.BorderColor3DBase = System.Drawing.Color.Black;
			appearance16.ForeColor = System.Drawing.Color.Firebrick;
			appearance16.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGDiciembre.Appearance = appearance16;
			this.txtFDGDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGDiciembre.Enabled = false;
			this.txtFDGDiciembre.FormatString = "#,##0.00";
			this.txtFDGDiciembre.Location = new System.Drawing.Point(1026, 688);
			this.txtFDGDiciembre.Name = "txtFDGDiciembre";
			this.txtFDGDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGDiciembre.PromptChar = ' ';
			this.txtFDGDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtFDGDiciembre.TabIndex = 1438;
			// 
			// txtFDGSeptiembre
			// 
			appearance17.BackColorDisabled = System.Drawing.Color.White;
			appearance17.BackColorDisabled2 = System.Drawing.Color.White;
			appearance17.BorderColor = System.Drawing.Color.Black;
			appearance17.BorderColor3DBase = System.Drawing.Color.Black;
			appearance17.ForeColor = System.Drawing.Color.Firebrick;
			appearance17.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGSeptiembre.Appearance = appearance17;
			this.txtFDGSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGSeptiembre.Enabled = false;
			this.txtFDGSeptiembre.FormatString = "#,##0.00";
			this.txtFDGSeptiembre.Location = new System.Drawing.Point(762, 688);
			this.txtFDGSeptiembre.Name = "txtFDGSeptiembre";
			this.txtFDGSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGSeptiembre.PromptChar = ' ';
			this.txtFDGSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtFDGSeptiembre.TabIndex = 1437;
			// 
			// txtFDGOctubre
			// 
			appearance18.BackColorDisabled = System.Drawing.Color.White;
			appearance18.BackColorDisabled2 = System.Drawing.Color.White;
			appearance18.BorderColor = System.Drawing.Color.Black;
			appearance18.BorderColor3DBase = System.Drawing.Color.Black;
			appearance18.ForeColor = System.Drawing.Color.Firebrick;
			appearance18.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGOctubre.Appearance = appearance18;
			this.txtFDGOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGOctubre.Enabled = false;
			this.txtFDGOctubre.FormatString = "#,##0.00";
			this.txtFDGOctubre.Location = new System.Drawing.Point(850, 688);
			this.txtFDGOctubre.Name = "txtFDGOctubre";
			this.txtFDGOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGOctubre.PromptChar = ' ';
			this.txtFDGOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtFDGOctubre.TabIndex = 1436;
			// 
			// txtFDGJulio
			// 
			appearance19.BackColorDisabled = System.Drawing.Color.White;
			appearance19.BackColorDisabled2 = System.Drawing.Color.White;
			appearance19.BorderColor = System.Drawing.Color.Black;
			appearance19.BorderColor3DBase = System.Drawing.Color.Black;
			appearance19.ForeColor = System.Drawing.Color.Firebrick;
			appearance19.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGJulio.Appearance = appearance19;
			this.txtFDGJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGJulio.Enabled = false;
			this.txtFDGJulio.FormatString = "#,##0.00";
			this.txtFDGJulio.Location = new System.Drawing.Point(586, 688);
			this.txtFDGJulio.Name = "txtFDGJulio";
			this.txtFDGJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGJulio.PromptChar = ' ';
			this.txtFDGJulio.Size = new System.Drawing.Size(80, 19);
			this.txtFDGJulio.TabIndex = 1435;
			// 
			// txtFDGAgosto
			// 
			appearance20.BackColorDisabled = System.Drawing.Color.White;
			appearance20.BackColorDisabled2 = System.Drawing.Color.White;
			appearance20.BorderColor = System.Drawing.Color.Black;
			appearance20.BorderColor3DBase = System.Drawing.Color.Black;
			appearance20.ForeColor = System.Drawing.Color.Firebrick;
			appearance20.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGAgosto.Appearance = appearance20;
			this.txtFDGAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGAgosto.Enabled = false;
			this.txtFDGAgosto.FormatString = "#,##0.00";
			this.txtFDGAgosto.Location = new System.Drawing.Point(674, 688);
			this.txtFDGAgosto.Name = "txtFDGAgosto";
			this.txtFDGAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGAgosto.PromptChar = ' ';
			this.txtFDGAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtFDGAgosto.TabIndex = 1434;
			// 
			// txtFDGMayo
			// 
			appearance21.BackColorDisabled = System.Drawing.Color.White;
			appearance21.BackColorDisabled2 = System.Drawing.Color.White;
			appearance21.BorderColor = System.Drawing.Color.Black;
			appearance21.BorderColor3DBase = System.Drawing.Color.Black;
			appearance21.ForeColor = System.Drawing.Color.Firebrick;
			appearance21.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGMayo.Appearance = appearance21;
			this.txtFDGMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGMayo.Enabled = false;
			this.txtFDGMayo.FormatString = "#,##0.00";
			this.txtFDGMayo.Location = new System.Drawing.Point(410, 688);
			this.txtFDGMayo.Name = "txtFDGMayo";
			this.txtFDGMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGMayo.PromptChar = ' ';
			this.txtFDGMayo.Size = new System.Drawing.Size(80, 19);
			this.txtFDGMayo.TabIndex = 1433;
			// 
			// txtFDGJunio
			// 
			appearance22.BackColorDisabled = System.Drawing.Color.White;
			appearance22.BackColorDisabled2 = System.Drawing.Color.White;
			appearance22.BorderColor = System.Drawing.Color.Black;
			appearance22.BorderColor3DBase = System.Drawing.Color.Black;
			appearance22.ForeColor = System.Drawing.Color.Firebrick;
			appearance22.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGJunio.Appearance = appearance22;
			this.txtFDGJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGJunio.Enabled = false;
			this.txtFDGJunio.FormatString = "#,##0.00";
			this.txtFDGJunio.Location = new System.Drawing.Point(498, 688);
			this.txtFDGJunio.Name = "txtFDGJunio";
			this.txtFDGJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGJunio.PromptChar = ' ';
			this.txtFDGJunio.Size = new System.Drawing.Size(80, 19);
			this.txtFDGJunio.TabIndex = 1432;
			// 
			// txtFDGMarzo
			// 
			appearance23.BackColorDisabled = System.Drawing.Color.White;
			appearance23.BackColorDisabled2 = System.Drawing.Color.White;
			appearance23.BorderColor = System.Drawing.Color.Black;
			appearance23.BorderColor3DBase = System.Drawing.Color.Black;
			appearance23.ForeColor = System.Drawing.Color.Firebrick;
			appearance23.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGMarzo.Appearance = appearance23;
			this.txtFDGMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGMarzo.Enabled = false;
			this.txtFDGMarzo.FormatString = "#,##0.00";
			this.txtFDGMarzo.Location = new System.Drawing.Point(234, 688);
			this.txtFDGMarzo.Name = "txtFDGMarzo";
			this.txtFDGMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGMarzo.PromptChar = ' ';
			this.txtFDGMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtFDGMarzo.TabIndex = 1431;
			// 
			// txtFDGAbril
			// 
			appearance24.BackColorDisabled = System.Drawing.Color.White;
			appearance24.BackColorDisabled2 = System.Drawing.Color.White;
			appearance24.BorderColor = System.Drawing.Color.Black;
			appearance24.BorderColor3DBase = System.Drawing.Color.Black;
			appearance24.ForeColor = System.Drawing.Color.Firebrick;
			appearance24.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGAbril.Appearance = appearance24;
			this.txtFDGAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGAbril.Enabled = false;
			this.txtFDGAbril.FormatString = "#,##0.00";
			this.txtFDGAbril.Location = new System.Drawing.Point(322, 688);
			this.txtFDGAbril.Name = "txtFDGAbril";
			this.txtFDGAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGAbril.PromptChar = ' ';
			this.txtFDGAbril.Size = new System.Drawing.Size(80, 19);
			this.txtFDGAbril.TabIndex = 1430;
			// 
			// txtFDGFebrero
			// 
			appearance25.BackColorDisabled = System.Drawing.Color.White;
			appearance25.BackColorDisabled2 = System.Drawing.Color.White;
			appearance25.BorderColor = System.Drawing.Color.Black;
			appearance25.BorderColor3DBase = System.Drawing.Color.Black;
			appearance25.ForeColor = System.Drawing.Color.Firebrick;
			appearance25.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGFebrero.Appearance = appearance25;
			this.txtFDGFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGFebrero.Enabled = false;
			this.txtFDGFebrero.FormatString = "#,##0.00";
			this.txtFDGFebrero.Location = new System.Drawing.Point(146, 688);
			this.txtFDGFebrero.Name = "txtFDGFebrero";
			this.txtFDGFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGFebrero.PromptChar = ' ';
			this.txtFDGFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtFDGFebrero.TabIndex = 1429;
			// 
			// txtFDGEnero
			// 
			appearance26.BackColorDisabled = System.Drawing.Color.White;
			appearance26.BackColorDisabled2 = System.Drawing.Color.White;
			appearance26.BorderColor = System.Drawing.Color.Black;
			appearance26.BorderColor3DBase = System.Drawing.Color.Black;
			appearance26.ForeColor = System.Drawing.Color.Firebrick;
			appearance26.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFDGEnero.Appearance = appearance26;
			this.txtFDGEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtFDGEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFDGEnero.Enabled = false;
			this.txtFDGEnero.FormatString = "#,##0.00";
			this.txtFDGEnero.Location = new System.Drawing.Point(56, 688);
			this.txtFDGEnero.Name = "txtFDGEnero";
			this.txtFDGEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFDGEnero.PromptChar = ' ';
			this.txtFDGEnero.Size = new System.Drawing.Size(80, 19);
			this.txtFDGEnero.TabIndex = 1428;
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label38.ForeColor = System.Drawing.Color.Black;
			this.label38.Location = new System.Drawing.Point(13, 664);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(28, 16);
			this.label38.TabIndex = 1427;
			this.label38.Text = "- PR";
			this.label38.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPrestamosTotal
			// 
			appearance27.BackColorDisabled = System.Drawing.Color.White;
			appearance27.BackColorDisabled2 = System.Drawing.Color.White;
			appearance27.BorderColor = System.Drawing.Color.Black;
			appearance27.BorderColor3DBase = System.Drawing.Color.Black;
			appearance27.FontData.BoldAsString = "True";
			appearance27.ForeColor = System.Drawing.Color.Firebrick;
			appearance27.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosTotal.Appearance = appearance27;
			this.txtPrestamosTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosTotal.Enabled = false;
			this.txtPrestamosTotal.FormatString = "#,##0.00";
			this.txtPrestamosTotal.Location = new System.Drawing.Point(1114, 664);
			this.txtPrestamosTotal.Name = "txtPrestamosTotal";
			this.txtPrestamosTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosTotal.PromptChar = ' ';
			this.txtPrestamosTotal.Size = new System.Drawing.Size(100, 19);
			this.txtPrestamosTotal.TabIndex = 1426;
			// 
			// txtPrestamosNoviembre
			// 
			appearance28.BackColorDisabled = System.Drawing.Color.White;
			appearance28.BackColorDisabled2 = System.Drawing.Color.White;
			appearance28.BorderColor = System.Drawing.Color.Black;
			appearance28.BorderColor3DBase = System.Drawing.Color.Black;
			appearance28.ForeColor = System.Drawing.Color.Firebrick;
			appearance28.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosNoviembre.Appearance = appearance28;
			this.txtPrestamosNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosNoviembre.Enabled = false;
			this.txtPrestamosNoviembre.FormatString = "#,##0.00";
			this.txtPrestamosNoviembre.Location = new System.Drawing.Point(938, 664);
			this.txtPrestamosNoviembre.Name = "txtPrestamosNoviembre";
			this.txtPrestamosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosNoviembre.PromptChar = ' ';
			this.txtPrestamosNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosNoviembre.TabIndex = 1425;
			// 
			// txtPrestamosDiciembre
			// 
			appearance29.BackColorDisabled = System.Drawing.Color.White;
			appearance29.BackColorDisabled2 = System.Drawing.Color.White;
			appearance29.BorderColor = System.Drawing.Color.Black;
			appearance29.BorderColor3DBase = System.Drawing.Color.Black;
			appearance29.ForeColor = System.Drawing.Color.Firebrick;
			appearance29.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosDiciembre.Appearance = appearance29;
			this.txtPrestamosDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosDiciembre.Enabled = false;
			this.txtPrestamosDiciembre.FormatString = "#,##0.00";
			this.txtPrestamosDiciembre.Location = new System.Drawing.Point(1026, 664);
			this.txtPrestamosDiciembre.Name = "txtPrestamosDiciembre";
			this.txtPrestamosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosDiciembre.PromptChar = ' ';
			this.txtPrestamosDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosDiciembre.TabIndex = 1424;
			// 
			// txtPrestamosSeptiembre
			// 
			appearance30.BackColorDisabled = System.Drawing.Color.White;
			appearance30.BackColorDisabled2 = System.Drawing.Color.White;
			appearance30.BorderColor = System.Drawing.Color.Black;
			appearance30.BorderColor3DBase = System.Drawing.Color.Black;
			appearance30.ForeColor = System.Drawing.Color.Firebrick;
			appearance30.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosSeptiembre.Appearance = appearance30;
			this.txtPrestamosSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosSeptiembre.Enabled = false;
			this.txtPrestamosSeptiembre.FormatString = "#,##0.00";
			this.txtPrestamosSeptiembre.Location = new System.Drawing.Point(762, 664);
			this.txtPrestamosSeptiembre.Name = "txtPrestamosSeptiembre";
			this.txtPrestamosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosSeptiembre.PromptChar = ' ';
			this.txtPrestamosSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosSeptiembre.TabIndex = 1423;
			// 
			// txtPrestamosOctubre
			// 
			appearance31.BackColorDisabled = System.Drawing.Color.White;
			appearance31.BackColorDisabled2 = System.Drawing.Color.White;
			appearance31.BorderColor = System.Drawing.Color.Black;
			appearance31.BorderColor3DBase = System.Drawing.Color.Black;
			appearance31.ForeColor = System.Drawing.Color.Firebrick;
			appearance31.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosOctubre.Appearance = appearance31;
			this.txtPrestamosOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosOctubre.Enabled = false;
			this.txtPrestamosOctubre.FormatString = "#,##0.00";
			this.txtPrestamosOctubre.Location = new System.Drawing.Point(850, 664);
			this.txtPrestamosOctubre.Name = "txtPrestamosOctubre";
			this.txtPrestamosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosOctubre.PromptChar = ' ';
			this.txtPrestamosOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosOctubre.TabIndex = 1422;
			// 
			// txtPrestamosJulio
			// 
			appearance32.BackColorDisabled = System.Drawing.Color.White;
			appearance32.BackColorDisabled2 = System.Drawing.Color.White;
			appearance32.BorderColor = System.Drawing.Color.Black;
			appearance32.BorderColor3DBase = System.Drawing.Color.Black;
			appearance32.ForeColor = System.Drawing.Color.Firebrick;
			appearance32.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosJulio.Appearance = appearance32;
			this.txtPrestamosJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosJulio.Enabled = false;
			this.txtPrestamosJulio.FormatString = "#,##0.00";
			this.txtPrestamosJulio.Location = new System.Drawing.Point(586, 664);
			this.txtPrestamosJulio.Name = "txtPrestamosJulio";
			this.txtPrestamosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosJulio.PromptChar = ' ';
			this.txtPrestamosJulio.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosJulio.TabIndex = 1421;
			// 
			// txtPrestamosAgosto
			// 
			appearance33.BackColorDisabled = System.Drawing.Color.White;
			appearance33.BackColorDisabled2 = System.Drawing.Color.White;
			appearance33.BorderColor = System.Drawing.Color.Black;
			appearance33.BorderColor3DBase = System.Drawing.Color.Black;
			appearance33.ForeColor = System.Drawing.Color.Firebrick;
			appearance33.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosAgosto.Appearance = appearance33;
			this.txtPrestamosAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosAgosto.Enabled = false;
			this.txtPrestamosAgosto.FormatString = "#,##0.00";
			this.txtPrestamosAgosto.Location = new System.Drawing.Point(674, 664);
			this.txtPrestamosAgosto.Name = "txtPrestamosAgosto";
			this.txtPrestamosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosAgosto.PromptChar = ' ';
			this.txtPrestamosAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosAgosto.TabIndex = 1420;
			// 
			// txtPrestamosMayo
			// 
			appearance34.BackColorDisabled = System.Drawing.Color.White;
			appearance34.BackColorDisabled2 = System.Drawing.Color.White;
			appearance34.BorderColor = System.Drawing.Color.Black;
			appearance34.BorderColor3DBase = System.Drawing.Color.Black;
			appearance34.ForeColor = System.Drawing.Color.Firebrick;
			appearance34.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosMayo.Appearance = appearance34;
			this.txtPrestamosMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosMayo.Enabled = false;
			this.txtPrestamosMayo.FormatString = "#,##0.00";
			this.txtPrestamosMayo.Location = new System.Drawing.Point(410, 664);
			this.txtPrestamosMayo.Name = "txtPrestamosMayo";
			this.txtPrestamosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosMayo.PromptChar = ' ';
			this.txtPrestamosMayo.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosMayo.TabIndex = 1419;
			// 
			// txtPrestamosJunio
			// 
			appearance35.BackColorDisabled = System.Drawing.Color.White;
			appearance35.BackColorDisabled2 = System.Drawing.Color.White;
			appearance35.BorderColor = System.Drawing.Color.Black;
			appearance35.BorderColor3DBase = System.Drawing.Color.Black;
			appearance35.ForeColor = System.Drawing.Color.Firebrick;
			appearance35.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosJunio.Appearance = appearance35;
			this.txtPrestamosJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosJunio.Enabled = false;
			this.txtPrestamosJunio.FormatString = "#,##0.00";
			this.txtPrestamosJunio.Location = new System.Drawing.Point(498, 664);
			this.txtPrestamosJunio.Name = "txtPrestamosJunio";
			this.txtPrestamosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosJunio.PromptChar = ' ';
			this.txtPrestamosJunio.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosJunio.TabIndex = 1418;
			// 
			// txtPrestamosMarzo
			// 
			appearance36.BackColorDisabled = System.Drawing.Color.White;
			appearance36.BackColorDisabled2 = System.Drawing.Color.White;
			appearance36.BorderColor = System.Drawing.Color.Black;
			appearance36.BorderColor3DBase = System.Drawing.Color.Black;
			appearance36.ForeColor = System.Drawing.Color.Firebrick;
			appearance36.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosMarzo.Appearance = appearance36;
			this.txtPrestamosMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosMarzo.Enabled = false;
			this.txtPrestamosMarzo.FormatString = "#,##0.00";
			this.txtPrestamosMarzo.Location = new System.Drawing.Point(234, 664);
			this.txtPrestamosMarzo.Name = "txtPrestamosMarzo";
			this.txtPrestamosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosMarzo.PromptChar = ' ';
			this.txtPrestamosMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosMarzo.TabIndex = 1417;
			// 
			// txtPrestamosAbril
			// 
			appearance37.BackColorDisabled = System.Drawing.Color.White;
			appearance37.BackColorDisabled2 = System.Drawing.Color.White;
			appearance37.BorderColor = System.Drawing.Color.Black;
			appearance37.BorderColor3DBase = System.Drawing.Color.Black;
			appearance37.ForeColor = System.Drawing.Color.Firebrick;
			appearance37.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosAbril.Appearance = appearance37;
			this.txtPrestamosAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosAbril.Enabled = false;
			this.txtPrestamosAbril.FormatString = "#,##0.00";
			this.txtPrestamosAbril.Location = new System.Drawing.Point(322, 664);
			this.txtPrestamosAbril.Name = "txtPrestamosAbril";
			this.txtPrestamosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosAbril.PromptChar = ' ';
			this.txtPrestamosAbril.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosAbril.TabIndex = 1416;
			// 
			// txtPrestamosFebrero
			// 
			appearance38.BackColorDisabled = System.Drawing.Color.White;
			appearance38.BackColorDisabled2 = System.Drawing.Color.White;
			appearance38.BorderColor = System.Drawing.Color.Black;
			appearance38.BorderColor3DBase = System.Drawing.Color.Black;
			appearance38.ForeColor = System.Drawing.Color.Firebrick;
			appearance38.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosFebrero.Appearance = appearance38;
			this.txtPrestamosFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosFebrero.Enabled = false;
			this.txtPrestamosFebrero.FormatString = "#,##0.00";
			this.txtPrestamosFebrero.Location = new System.Drawing.Point(146, 664);
			this.txtPrestamosFebrero.Name = "txtPrestamosFebrero";
			this.txtPrestamosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosFebrero.PromptChar = ' ';
			this.txtPrestamosFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosFebrero.TabIndex = 1415;
			// 
			// txtPrestamosEnero
			// 
			appearance39.BackColorDisabled = System.Drawing.Color.White;
			appearance39.BackColorDisabled2 = System.Drawing.Color.White;
			appearance39.BorderColor = System.Drawing.Color.Black;
			appearance39.BorderColor3DBase = System.Drawing.Color.Black;
			appearance39.ForeColor = System.Drawing.Color.Firebrick;
			appearance39.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtPrestamosEnero.Appearance = appearance39;
			this.txtPrestamosEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtPrestamosEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrestamosEnero.Enabled = false;
			this.txtPrestamosEnero.FormatString = "#,##0.00";
			this.txtPrestamosEnero.Location = new System.Drawing.Point(58, 664);
			this.txtPrestamosEnero.Name = "txtPrestamosEnero";
			this.txtPrestamosEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrestamosEnero.PromptChar = ' ';
			this.txtPrestamosEnero.Size = new System.Drawing.Size(80, 19);
			this.txtPrestamosEnero.TabIndex = 1414;
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label37.ForeColor = System.Drawing.Color.Firebrick;
			this.label37.Location = new System.Drawing.Point(10, 352);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(23, 16);
			this.label37.TabIndex = 1413;
			this.label37.Text = "+ C";
			this.label37.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label36.ForeColor = System.Drawing.Color.Firebrick;
			this.label36.Location = new System.Drawing.Point(9, 336);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(24, 16);
			this.label36.TabIndex = 1412;
			this.label36.Text = "= M";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label35.ForeColor = System.Drawing.Color.Firebrick;
			this.label35.Location = new System.Drawing.Point(10, 192);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(22, 16);
			this.label35.TabIndex = 1411;
			this.label35.Text = "= V";
			this.label35.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridGastos
			// 
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource1;
			appearance40.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance40;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.Caption = "Año";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 137;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Gasto";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 97;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance41;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 81;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance42;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 81;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance43;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 78;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance44;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Width = 78;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance45;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.Width = 78;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance46;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 12;
			ultraGridColumn9.Width = 78;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance47;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 14;
			ultraGridColumn10.Width = 78;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance48;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 16;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance49;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 18;
			ultraGridColumn12.Width = 82;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance50;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 20;
			ultraGridColumn13.Width = 82;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance51;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 22;
			ultraGridColumn14.Width = 82;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance52;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 24;
			ultraGridColumn15.Width = 82;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance53;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 26;
			ultraGridColumn16.Width = 82;
			ultraGridColumn17.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn17.Header.Caption = "...";
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn17.Width = 20;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn18.Header.Caption = "...";
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 20;
			ultraGridColumn19.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn19.Header.Caption = "...";
			ultraGridColumn19.Header.VisiblePosition = 9;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 20;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn20.Header.Caption = "...";
			ultraGridColumn20.Header.VisiblePosition = 11;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn20.Width = 20;
			ultraGridColumn21.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn21.Header.Caption = "...";
			ultraGridColumn21.Header.VisiblePosition = 13;
			ultraGridColumn21.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn21.Width = 20;
			ultraGridColumn22.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn22.Header.Caption = "...";
			ultraGridColumn22.Header.VisiblePosition = 15;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn22.Width = 20;
			ultraGridColumn23.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn23.Header.Caption = "...";
			ultraGridColumn23.Header.VisiblePosition = 17;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn23.Width = 20;
			ultraGridColumn24.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn24.Header.Caption = "...";
			ultraGridColumn24.Header.VisiblePosition = 19;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn24.Width = 20;
			ultraGridColumn25.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn25.Header.Caption = "...";
			ultraGridColumn25.Header.VisiblePosition = 21;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn25.Width = 20;
			ultraGridColumn26.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn26.Header.Caption = "...";
			ultraGridColumn26.Header.VisiblePosition = 23;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn26.Width = 20;
			ultraGridColumn27.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn27.Header.Caption = "...";
			ultraGridColumn27.Header.VisiblePosition = 25;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn27.Width = 20;
			ultraGridColumn28.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn28.Header.Caption = "...";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn28.Width = 20;
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
																										 ultraGridColumn28});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
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
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance54.ForeColor = System.Drawing.Color.Black;
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance54;
			this.uGridGastos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGastos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance55.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance55;
			appearance56.ForeColor = System.Drawing.Color.Black;
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.CellAppearance = appearance56;
			this.uGridGastos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance57.FontData.BoldAsString = "True";
			appearance57.FontData.Name = "Arial";
			appearance57.FontData.SizeInPoints = 8F;
			appearance57.ForeColor = System.Drawing.Color.White;
			appearance57.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance57;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance58.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance58.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance58;
			appearance59.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance59.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance59;
			appearance60.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance60.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance60.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance60;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(3, 400);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(1215, 136);
			this.uGridGastos.TabIndex = 1270;
			this.uGridGastos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_ClickCellButton);
			this.uGridGastos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGastos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
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
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
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
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label33.ForeColor = System.Drawing.Color.Black;
			this.label33.Location = new System.Drawing.Point(5, 592);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(44, 16);
			this.label33.TabIndex = 1410;
			this.label33.Text = "+ NC M";
			this.label33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label34.ForeColor = System.Drawing.Color.Black;
			this.label34.Location = new System.Drawing.Point(6, 568);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(43, 16);
			this.label34.TabIndex = 1409;
			this.label34.Text = "+ G NC";
			this.label34.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtGNCTotal
			// 
			appearance61.BackColorDisabled = System.Drawing.Color.White;
			appearance61.BackColorDisabled2 = System.Drawing.Color.White;
			appearance61.BorderColor = System.Drawing.Color.Black;
			appearance61.BorderColor3DBase = System.Drawing.Color.Black;
			appearance61.FontData.BoldAsString = "True";
			appearance61.ForeColor = System.Drawing.Color.Black;
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCTotal.Appearance = appearance61;
			this.txtGNCTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCTotal.Enabled = false;
			this.txtGNCTotal.FormatString = "#,##0.00";
			this.txtGNCTotal.Location = new System.Drawing.Point(1114, 568);
			this.txtGNCTotal.Name = "txtGNCTotal";
			this.txtGNCTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCTotal.PromptChar = ' ';
			this.txtGNCTotal.ReadOnly = true;
			this.txtGNCTotal.Size = new System.Drawing.Size(100, 19);
			this.txtGNCTotal.TabIndex = 1408;
			this.txtGNCTotal.DoubleClick += new System.EventHandler(this.txtGNCTotal_DoubleClick);
			// 
			// txtGNCNoviembre
			// 
			appearance62.BackColorDisabled = System.Drawing.Color.White;
			appearance62.BackColorDisabled2 = System.Drawing.Color.White;
			appearance62.BorderColor = System.Drawing.Color.Black;
			appearance62.BorderColor3DBase = System.Drawing.Color.Black;
			appearance62.ForeColor = System.Drawing.Color.Black;
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCNoviembre.Appearance = appearance62;
			this.txtGNCNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCNoviembre.FormatString = "#,##0.00";
			this.txtGNCNoviembre.Location = new System.Drawing.Point(938, 568);
			this.txtGNCNoviembre.Name = "txtGNCNoviembre";
			this.txtGNCNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCNoviembre.PromptChar = ' ';
			this.txtGNCNoviembre.ReadOnly = true;
			this.txtGNCNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtGNCNoviembre.TabIndex = 1407;
			this.txtGNCNoviembre.DoubleClick += new System.EventHandler(this.txtGNCNoviembre_DoubleClick);
			// 
			// txtGNCDiciembre
			// 
			appearance63.BackColorDisabled = System.Drawing.Color.White;
			appearance63.BackColorDisabled2 = System.Drawing.Color.White;
			appearance63.BorderColor = System.Drawing.Color.Black;
			appearance63.BorderColor3DBase = System.Drawing.Color.Black;
			appearance63.ForeColor = System.Drawing.Color.Black;
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCDiciembre.Appearance = appearance63;
			this.txtGNCDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCDiciembre.FormatString = "#,##0.00";
			this.txtGNCDiciembre.Location = new System.Drawing.Point(1026, 568);
			this.txtGNCDiciembre.Name = "txtGNCDiciembre";
			this.txtGNCDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCDiciembre.PromptChar = ' ';
			this.txtGNCDiciembre.ReadOnly = true;
			this.txtGNCDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtGNCDiciembre.TabIndex = 1406;
			this.txtGNCDiciembre.DoubleClick += new System.EventHandler(this.txtGNCDiciembre_DoubleClick);
			// 
			// txtGNCSeptiembre
			// 
			appearance64.BackColorDisabled = System.Drawing.Color.White;
			appearance64.BackColorDisabled2 = System.Drawing.Color.White;
			appearance64.BorderColor = System.Drawing.Color.Black;
			appearance64.BorderColor3DBase = System.Drawing.Color.Black;
			appearance64.ForeColor = System.Drawing.Color.Black;
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCSeptiembre.Appearance = appearance64;
			this.txtGNCSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCSeptiembre.FormatString = "#,##0.00";
			this.txtGNCSeptiembre.Location = new System.Drawing.Point(762, 568);
			this.txtGNCSeptiembre.Name = "txtGNCSeptiembre";
			this.txtGNCSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCSeptiembre.PromptChar = ' ';
			this.txtGNCSeptiembre.ReadOnly = true;
			this.txtGNCSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtGNCSeptiembre.TabIndex = 1405;
			this.txtGNCSeptiembre.DoubleClick += new System.EventHandler(this.txtGNCSeptiembre_DoubleClick);
			// 
			// txtGNCOctubre
			// 
			appearance65.BackColorDisabled = System.Drawing.Color.White;
			appearance65.BackColorDisabled2 = System.Drawing.Color.White;
			appearance65.BorderColor = System.Drawing.Color.Black;
			appearance65.BorderColor3DBase = System.Drawing.Color.Black;
			appearance65.ForeColor = System.Drawing.Color.Black;
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCOctubre.Appearance = appearance65;
			this.txtGNCOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCOctubre.FormatString = "#,##0.00";
			this.txtGNCOctubre.Location = new System.Drawing.Point(850, 568);
			this.txtGNCOctubre.Name = "txtGNCOctubre";
			this.txtGNCOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCOctubre.PromptChar = ' ';
			this.txtGNCOctubre.ReadOnly = true;
			this.txtGNCOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtGNCOctubre.TabIndex = 1404;
			this.txtGNCOctubre.DoubleClick += new System.EventHandler(this.txtGNCOctubre_DoubleClick);
			// 
			// txtGNCJulio
			// 
			appearance66.BackColorDisabled = System.Drawing.Color.White;
			appearance66.BackColorDisabled2 = System.Drawing.Color.White;
			appearance66.BorderColor = System.Drawing.Color.Black;
			appearance66.BorderColor3DBase = System.Drawing.Color.Black;
			appearance66.ForeColor = System.Drawing.Color.Black;
			appearance66.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCJulio.Appearance = appearance66;
			this.txtGNCJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCJulio.FormatString = "#,##0.00";
			this.txtGNCJulio.Location = new System.Drawing.Point(586, 568);
			this.txtGNCJulio.Name = "txtGNCJulio";
			this.txtGNCJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCJulio.PromptChar = ' ';
			this.txtGNCJulio.ReadOnly = true;
			this.txtGNCJulio.Size = new System.Drawing.Size(80, 19);
			this.txtGNCJulio.TabIndex = 1403;
			this.txtGNCJulio.DoubleClick += new System.EventHandler(this.txtGNCJulio_DoubleClick);
			this.txtGNCJulio.ValueChanged += new System.EventHandler(this.txtGNCJulio_ValueChanged);
			// 
			// txtGNCAgosto
			// 
			appearance67.BackColorDisabled = System.Drawing.Color.White;
			appearance67.BackColorDisabled2 = System.Drawing.Color.White;
			appearance67.BorderColor = System.Drawing.Color.Black;
			appearance67.BorderColor3DBase = System.Drawing.Color.Black;
			appearance67.ForeColor = System.Drawing.Color.Black;
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCAgosto.Appearance = appearance67;
			this.txtGNCAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCAgosto.FormatString = "#,##0.00";
			this.txtGNCAgosto.Location = new System.Drawing.Point(674, 568);
			this.txtGNCAgosto.Name = "txtGNCAgosto";
			this.txtGNCAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCAgosto.PromptChar = ' ';
			this.txtGNCAgosto.ReadOnly = true;
			this.txtGNCAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtGNCAgosto.TabIndex = 1402;
			this.txtGNCAgosto.DoubleClick += new System.EventHandler(this.txtGNCAgosto_DoubleClick);
			this.txtGNCAgosto.ValueChanged += new System.EventHandler(this.txtGNCAgosto_ValueChanged);
			// 
			// txtGNCMayo
			// 
			appearance68.BackColorDisabled = System.Drawing.Color.White;
			appearance68.BackColorDisabled2 = System.Drawing.Color.White;
			appearance68.BorderColor = System.Drawing.Color.Black;
			appearance68.BorderColor3DBase = System.Drawing.Color.Black;
			appearance68.ForeColor = System.Drawing.Color.Black;
			appearance68.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCMayo.Appearance = appearance68;
			this.txtGNCMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCMayo.FormatString = "#,##0.00";
			this.txtGNCMayo.Location = new System.Drawing.Point(410, 568);
			this.txtGNCMayo.Name = "txtGNCMayo";
			this.txtGNCMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCMayo.PromptChar = ' ';
			this.txtGNCMayo.ReadOnly = true;
			this.txtGNCMayo.Size = new System.Drawing.Size(80, 19);
			this.txtGNCMayo.TabIndex = 1401;
			this.txtGNCMayo.DoubleClick += new System.EventHandler(this.txtGNCMayo_DoubleClick);
			// 
			// txtGNCJunio
			// 
			appearance69.BackColorDisabled = System.Drawing.Color.White;
			appearance69.BackColorDisabled2 = System.Drawing.Color.White;
			appearance69.BorderColor = System.Drawing.Color.Black;
			appearance69.BorderColor3DBase = System.Drawing.Color.Black;
			appearance69.ForeColor = System.Drawing.Color.Black;
			appearance69.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCJunio.Appearance = appearance69;
			this.txtGNCJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCJunio.FormatString = "#,##0.00";
			this.txtGNCJunio.Location = new System.Drawing.Point(498, 568);
			this.txtGNCJunio.Name = "txtGNCJunio";
			this.txtGNCJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCJunio.PromptChar = ' ';
			this.txtGNCJunio.ReadOnly = true;
			this.txtGNCJunio.Size = new System.Drawing.Size(80, 19);
			this.txtGNCJunio.TabIndex = 1400;
			this.txtGNCJunio.DoubleClick += new System.EventHandler(this.txtGNCJunio_DoubleClick);
			// 
			// txtGNCMarzo
			// 
			appearance70.BackColorDisabled = System.Drawing.Color.White;
			appearance70.BackColorDisabled2 = System.Drawing.Color.White;
			appearance70.BorderColor = System.Drawing.Color.Black;
			appearance70.BorderColor3DBase = System.Drawing.Color.Black;
			appearance70.ForeColor = System.Drawing.Color.Black;
			appearance70.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCMarzo.Appearance = appearance70;
			this.txtGNCMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCMarzo.FormatString = "#,##0.00";
			this.txtGNCMarzo.Location = new System.Drawing.Point(234, 568);
			this.txtGNCMarzo.Name = "txtGNCMarzo";
			this.txtGNCMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCMarzo.PromptChar = ' ';
			this.txtGNCMarzo.ReadOnly = true;
			this.txtGNCMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtGNCMarzo.TabIndex = 1399;
			this.txtGNCMarzo.DoubleClick += new System.EventHandler(this.txtGNCMarzo_DoubleClick);
			// 
			// txtGNCAbril
			// 
			appearance71.BackColorDisabled = System.Drawing.Color.White;
			appearance71.BackColorDisabled2 = System.Drawing.Color.White;
			appearance71.BorderColor = System.Drawing.Color.Black;
			appearance71.BorderColor3DBase = System.Drawing.Color.Black;
			appearance71.ForeColor = System.Drawing.Color.Black;
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCAbril.Appearance = appearance71;
			this.txtGNCAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCAbril.FormatString = "#,##0.00";
			this.txtGNCAbril.Location = new System.Drawing.Point(322, 568);
			this.txtGNCAbril.Name = "txtGNCAbril";
			this.txtGNCAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCAbril.PromptChar = ' ';
			this.txtGNCAbril.ReadOnly = true;
			this.txtGNCAbril.Size = new System.Drawing.Size(80, 19);
			this.txtGNCAbril.TabIndex = 1398;
			this.txtGNCAbril.DoubleClick += new System.EventHandler(this.txtGNCAbril_DoubleClick);
			// 
			// txtGNCFebrero
			// 
			appearance72.BackColorDisabled = System.Drawing.Color.White;
			appearance72.BackColorDisabled2 = System.Drawing.Color.White;
			appearance72.BorderColor = System.Drawing.Color.Black;
			appearance72.BorderColor3DBase = System.Drawing.Color.Black;
			appearance72.ForeColor = System.Drawing.Color.Black;
			appearance72.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCFebrero.Appearance = appearance72;
			this.txtGNCFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCFebrero.FormatString = "#,##0.00";
			this.txtGNCFebrero.Location = new System.Drawing.Point(146, 568);
			this.txtGNCFebrero.Name = "txtGNCFebrero";
			this.txtGNCFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCFebrero.PromptChar = ' ';
			this.txtGNCFebrero.ReadOnly = true;
			this.txtGNCFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtGNCFebrero.TabIndex = 1397;
			this.txtGNCFebrero.DoubleClick += new System.EventHandler(this.txtGNCFebrero_DoubleClick);
			// 
			// txtGNCEnero
			// 
			appearance73.BackColorDisabled = System.Drawing.Color.White;
			appearance73.BackColorDisabled2 = System.Drawing.Color.White;
			appearance73.BorderColor = System.Drawing.Color.Black;
			appearance73.BorderColor3DBase = System.Drawing.Color.Black;
			appearance73.ForeColor = System.Drawing.Color.Black;
			appearance73.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGNCEnero.Appearance = appearance73;
			this.txtGNCEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGNCEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGNCEnero.FormatString = "#,##0.00";
			this.txtGNCEnero.Location = new System.Drawing.Point(58, 568);
			this.txtGNCEnero.Name = "txtGNCEnero";
			this.txtGNCEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGNCEnero.PromptChar = ' ';
			this.txtGNCEnero.ReadOnly = true;
			this.txtGNCEnero.Size = new System.Drawing.Size(80, 19);
			this.txtGNCEnero.TabIndex = 1396;
			this.toolTip1.SetToolTip(this.txtGNCEnero, "Doble Click para ver el detalle o ingresar nuevos documentos");
			this.txtGNCEnero.DoubleClick += new System.EventHandler(this.txtGNCEnero_DoubleClick);
			// 
			// txtMNCTotal
			// 
			appearance74.BackColorDisabled = System.Drawing.Color.White;
			appearance74.BackColorDisabled2 = System.Drawing.Color.White;
			appearance74.BorderColor = System.Drawing.Color.Black;
			appearance74.BorderColor3DBase = System.Drawing.Color.Black;
			appearance74.FontData.BoldAsString = "True";
			appearance74.ForeColor = System.Drawing.Color.Black;
			appearance74.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCTotal.Appearance = appearance74;
			this.txtMNCTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCTotal.Enabled = false;
			this.txtMNCTotal.FormatString = "#,##0.00";
			this.txtMNCTotal.Location = new System.Drawing.Point(1114, 592);
			this.txtMNCTotal.Name = "txtMNCTotal";
			this.txtMNCTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCTotal.PromptChar = ' ';
			this.txtMNCTotal.ReadOnly = true;
			this.txtMNCTotal.Size = new System.Drawing.Size(100, 19);
			this.txtMNCTotal.TabIndex = 1395;
			this.txtMNCTotal.DoubleClick += new System.EventHandler(this.txtMNCTotal_DoubleClick);
			// 
			// txtMNCNoviembre
			// 
			appearance75.BackColorDisabled = System.Drawing.Color.White;
			appearance75.BackColorDisabled2 = System.Drawing.Color.White;
			appearance75.BorderColor = System.Drawing.Color.Black;
			appearance75.BorderColor3DBase = System.Drawing.Color.Black;
			appearance75.ForeColor = System.Drawing.Color.Black;
			appearance75.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCNoviembre.Appearance = appearance75;
			this.txtMNCNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCNoviembre.FormatString = "#,##0.00";
			this.txtMNCNoviembre.Location = new System.Drawing.Point(938, 592);
			this.txtMNCNoviembre.Name = "txtMNCNoviembre";
			this.txtMNCNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCNoviembre.PromptChar = ' ';
			this.txtMNCNoviembre.ReadOnly = true;
			this.txtMNCNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtMNCNoviembre.TabIndex = 1394;
			this.txtMNCNoviembre.DoubleClick += new System.EventHandler(this.txtMNCNoviembre_DoubleClick);
			// 
			// txtMNCDiciembre
			// 
			appearance76.BackColorDisabled = System.Drawing.Color.White;
			appearance76.BackColorDisabled2 = System.Drawing.Color.White;
			appearance76.BorderColor = System.Drawing.Color.Black;
			appearance76.BorderColor3DBase = System.Drawing.Color.Black;
			appearance76.ForeColor = System.Drawing.Color.Black;
			appearance76.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCDiciembre.Appearance = appearance76;
			this.txtMNCDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCDiciembre.FormatString = "#,##0.00";
			this.txtMNCDiciembre.Location = new System.Drawing.Point(1026, 592);
			this.txtMNCDiciembre.Name = "txtMNCDiciembre";
			this.txtMNCDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCDiciembre.PromptChar = ' ';
			this.txtMNCDiciembre.ReadOnly = true;
			this.txtMNCDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtMNCDiciembre.TabIndex = 1393;
			this.txtMNCDiciembre.DoubleClick += new System.EventHandler(this.txtMNCDiciembre_DoubleClick);
			// 
			// txtMNCSeptiembre
			// 
			appearance77.BackColorDisabled = System.Drawing.Color.White;
			appearance77.BackColorDisabled2 = System.Drawing.Color.White;
			appearance77.BorderColor = System.Drawing.Color.Black;
			appearance77.BorderColor3DBase = System.Drawing.Color.Black;
			appearance77.ForeColor = System.Drawing.Color.Black;
			appearance77.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCSeptiembre.Appearance = appearance77;
			this.txtMNCSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCSeptiembre.FormatString = "#,##0.00";
			this.txtMNCSeptiembre.Location = new System.Drawing.Point(762, 592);
			this.txtMNCSeptiembre.Name = "txtMNCSeptiembre";
			this.txtMNCSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCSeptiembre.PromptChar = ' ';
			this.txtMNCSeptiembre.ReadOnly = true;
			this.txtMNCSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtMNCSeptiembre.TabIndex = 1392;
			this.txtMNCSeptiembre.DoubleClick += new System.EventHandler(this.txtMNCSeptiembre_DoubleClick);
			// 
			// txtMNCOctubre
			// 
			appearance78.BackColorDisabled = System.Drawing.Color.White;
			appearance78.BackColorDisabled2 = System.Drawing.Color.White;
			appearance78.BorderColor = System.Drawing.Color.Black;
			appearance78.BorderColor3DBase = System.Drawing.Color.Black;
			appearance78.ForeColor = System.Drawing.Color.Black;
			appearance78.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCOctubre.Appearance = appearance78;
			this.txtMNCOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCOctubre.FormatString = "#,##0.00";
			this.txtMNCOctubre.Location = new System.Drawing.Point(850, 592);
			this.txtMNCOctubre.Name = "txtMNCOctubre";
			this.txtMNCOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCOctubre.PromptChar = ' ';
			this.txtMNCOctubre.ReadOnly = true;
			this.txtMNCOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtMNCOctubre.TabIndex = 1391;
			this.txtMNCOctubre.DoubleClick += new System.EventHandler(this.txtMNCOctubre_DoubleClick);
			// 
			// txtMNCJulio
			// 
			appearance79.BackColorDisabled = System.Drawing.Color.White;
			appearance79.BackColorDisabled2 = System.Drawing.Color.White;
			appearance79.BorderColor = System.Drawing.Color.Black;
			appearance79.BorderColor3DBase = System.Drawing.Color.Black;
			appearance79.ForeColor = System.Drawing.Color.Black;
			appearance79.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCJulio.Appearance = appearance79;
			this.txtMNCJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCJulio.FormatString = "#,##0.00";
			this.txtMNCJulio.Location = new System.Drawing.Point(586, 592);
			this.txtMNCJulio.Name = "txtMNCJulio";
			this.txtMNCJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCJulio.PromptChar = ' ';
			this.txtMNCJulio.ReadOnly = true;
			this.txtMNCJulio.Size = new System.Drawing.Size(80, 19);
			this.txtMNCJulio.TabIndex = 1390;
			this.txtMNCJulio.DoubleClick += new System.EventHandler(this.txtMNCJulio_DoubleClick);
			// 
			// txtMNCAgosto
			// 
			appearance80.BackColorDisabled = System.Drawing.Color.White;
			appearance80.BackColorDisabled2 = System.Drawing.Color.White;
			appearance80.BorderColor = System.Drawing.Color.Black;
			appearance80.BorderColor3DBase = System.Drawing.Color.Black;
			appearance80.ForeColor = System.Drawing.Color.Black;
			appearance80.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCAgosto.Appearance = appearance80;
			this.txtMNCAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCAgosto.FormatString = "#,##0.00";
			this.txtMNCAgosto.Location = new System.Drawing.Point(674, 592);
			this.txtMNCAgosto.Name = "txtMNCAgosto";
			this.txtMNCAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCAgosto.PromptChar = ' ';
			this.txtMNCAgosto.ReadOnly = true;
			this.txtMNCAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtMNCAgosto.TabIndex = 1389;
			this.txtMNCAgosto.DoubleClick += new System.EventHandler(this.txtMNCAgosto_DoubleClick);
			// 
			// txtMNCMayo
			// 
			appearance81.BackColorDisabled = System.Drawing.Color.White;
			appearance81.BackColorDisabled2 = System.Drawing.Color.White;
			appearance81.BorderColor = System.Drawing.Color.Black;
			appearance81.BorderColor3DBase = System.Drawing.Color.Black;
			appearance81.ForeColor = System.Drawing.Color.Black;
			appearance81.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCMayo.Appearance = appearance81;
			this.txtMNCMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCMayo.FormatString = "#,##0.00";
			this.txtMNCMayo.Location = new System.Drawing.Point(410, 592);
			this.txtMNCMayo.Name = "txtMNCMayo";
			this.txtMNCMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCMayo.PromptChar = ' ';
			this.txtMNCMayo.ReadOnly = true;
			this.txtMNCMayo.Size = new System.Drawing.Size(80, 19);
			this.txtMNCMayo.TabIndex = 1388;
			this.txtMNCMayo.DoubleClick += new System.EventHandler(this.txtMNCMayo_DoubleClick);
			// 
			// txtMNCJunio
			// 
			appearance82.BackColorDisabled = System.Drawing.Color.White;
			appearance82.BackColorDisabled2 = System.Drawing.Color.White;
			appearance82.BorderColor = System.Drawing.Color.Black;
			appearance82.BorderColor3DBase = System.Drawing.Color.Black;
			appearance82.ForeColor = System.Drawing.Color.Black;
			appearance82.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCJunio.Appearance = appearance82;
			this.txtMNCJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCJunio.FormatString = "#,##0.00";
			this.txtMNCJunio.Location = new System.Drawing.Point(498, 592);
			this.txtMNCJunio.Name = "txtMNCJunio";
			this.txtMNCJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCJunio.PromptChar = ' ';
			this.txtMNCJunio.ReadOnly = true;
			this.txtMNCJunio.Size = new System.Drawing.Size(80, 19);
			this.txtMNCJunio.TabIndex = 1387;
			this.txtMNCJunio.DoubleClick += new System.EventHandler(this.txtMNCJunio_DoubleClick);
			// 
			// txtMNCMarzo
			// 
			appearance83.BackColorDisabled = System.Drawing.Color.White;
			appearance83.BackColorDisabled2 = System.Drawing.Color.White;
			appearance83.BorderColor = System.Drawing.Color.Black;
			appearance83.BorderColor3DBase = System.Drawing.Color.Black;
			appearance83.ForeColor = System.Drawing.Color.Black;
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCMarzo.Appearance = appearance83;
			this.txtMNCMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCMarzo.FormatString = "#,##0.00";
			this.txtMNCMarzo.Location = new System.Drawing.Point(234, 592);
			this.txtMNCMarzo.Name = "txtMNCMarzo";
			this.txtMNCMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCMarzo.PromptChar = ' ';
			this.txtMNCMarzo.ReadOnly = true;
			this.txtMNCMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtMNCMarzo.TabIndex = 1386;
			this.txtMNCMarzo.DoubleClick += new System.EventHandler(this.txtMNCMarzo_DoubleClick);
			// 
			// txtMNCAbril
			// 
			appearance84.BackColorDisabled = System.Drawing.Color.White;
			appearance84.BackColorDisabled2 = System.Drawing.Color.White;
			appearance84.BorderColor = System.Drawing.Color.Black;
			appearance84.BorderColor3DBase = System.Drawing.Color.Black;
			appearance84.ForeColor = System.Drawing.Color.Black;
			appearance84.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCAbril.Appearance = appearance84;
			this.txtMNCAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCAbril.FormatString = "#,##0.00";
			this.txtMNCAbril.Location = new System.Drawing.Point(322, 592);
			this.txtMNCAbril.Name = "txtMNCAbril";
			this.txtMNCAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCAbril.PromptChar = ' ';
			this.txtMNCAbril.ReadOnly = true;
			this.txtMNCAbril.Size = new System.Drawing.Size(80, 19);
			this.txtMNCAbril.TabIndex = 1385;
			this.txtMNCAbril.DoubleClick += new System.EventHandler(this.txtMNCAbril_DoubleClick);
			// 
			// txtMNCFebrero
			// 
			appearance85.BackColorDisabled = System.Drawing.Color.White;
			appearance85.BackColorDisabled2 = System.Drawing.Color.White;
			appearance85.BorderColor = System.Drawing.Color.Black;
			appearance85.BorderColor3DBase = System.Drawing.Color.Black;
			appearance85.ForeColor = System.Drawing.Color.Black;
			appearance85.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCFebrero.Appearance = appearance85;
			this.txtMNCFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCFebrero.FormatString = "#,##0.00";
			this.txtMNCFebrero.Location = new System.Drawing.Point(146, 592);
			this.txtMNCFebrero.Name = "txtMNCFebrero";
			this.txtMNCFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCFebrero.PromptChar = ' ';
			this.txtMNCFebrero.ReadOnly = true;
			this.txtMNCFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtMNCFebrero.TabIndex = 1384;
			this.txtMNCFebrero.DoubleClick += new System.EventHandler(this.txtMNCFebrero_DoubleClick);
			// 
			// txtMNCEnero
			// 
			appearance86.BackColorDisabled = System.Drawing.Color.White;
			appearance86.BackColorDisabled2 = System.Drawing.Color.White;
			appearance86.BorderColor = System.Drawing.Color.Black;
			appearance86.BorderColor3DBase = System.Drawing.Color.Black;
			appearance86.ForeColor = System.Drawing.Color.Black;
			appearance86.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMNCEnero.Appearance = appearance86;
			this.txtMNCEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMNCEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMNCEnero.FormatString = "#,##0.00";
			this.txtMNCEnero.Location = new System.Drawing.Point(58, 592);
			this.txtMNCEnero.Name = "txtMNCEnero";
			this.txtMNCEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMNCEnero.PromptChar = ' ';
			this.txtMNCEnero.ReadOnly = true;
			this.txtMNCEnero.Size = new System.Drawing.Size(80, 19);
			this.txtMNCEnero.TabIndex = 1383;
			this.txtMNCEnero.DoubleClick += new System.EventHandler(this.txtMNCEnero_DoubleClick);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.Location = new System.Drawing.Point(1152, 5);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(30, 23);
			this.btnExcel.TabIndex = 1382;
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label32.ForeColor = System.Drawing.Color.Red;
			this.label32.Location = new System.Drawing.Point(2, 496);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(82, 16);
			this.label32.TabIndex = 1381;
			this.label32.Text = "RET MES ANT";
			this.label32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label32.Visible = false;
			// 
			// txtSMAGastosRetirosDiciembre
			// 
			appearance87.BackColorDisabled = System.Drawing.Color.White;
			appearance87.BackColorDisabled2 = System.Drawing.Color.White;
			appearance87.BorderColor = System.Drawing.Color.Black;
			appearance87.BorderColor3DBase = System.Drawing.Color.Black;
			appearance87.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			appearance87.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.txtSMAGastosRetirosDiciembre.Appearance = appearance87;
			this.txtSMAGastosRetirosDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSMAGastosRetirosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSMAGastosRetirosDiciembre.Enabled = false;
			this.txtSMAGastosRetirosDiciembre.FormatString = "#,##0.00";
			this.txtSMAGastosRetirosDiciembre.Location = new System.Drawing.Point(986, 504);
			this.txtSMAGastosRetirosDiciembre.Name = "txtSMAGastosRetirosDiciembre";
			this.txtSMAGastosRetirosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSMAGastosRetirosDiciembre.PromptChar = ' ';
			this.txtSMAGastosRetirosDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtSMAGastosRetirosDiciembre.TabIndex = 1380;
			this.txtSMAGastosRetirosDiciembre.Visible = false;
			// 
			// txtSMAGastosRetirosNoviembre
			// 
			appearance88.BackColorDisabled = System.Drawing.Color.White;
			appearance88.BackColorDisabled2 = System.Drawing.Color.White;
			appearance88.BorderColor = System.Drawing.Color.Black;
			appearance88.BorderColor3DBase = System.Drawing.Color.Black;
			appearance88.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			appearance88.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.txtSMAGastosRetirosNoviembre.Appearance = appearance88;
			this.txtSMAGastosRetirosNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSMAGastosRetirosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSMAGastosRetirosNoviembre.Enabled = false;
			this.txtSMAGastosRetirosNoviembre.FormatString = "#,##0.00";
			this.txtSMAGastosRetirosNoviembre.Location = new System.Drawing.Point(898, 504);
			this.txtSMAGastosRetirosNoviembre.Name = "txtSMAGastosRetirosNoviembre";
			this.txtSMAGastosRetirosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSMAGastosRetirosNoviembre.PromptChar = ' ';
			this.txtSMAGastosRetirosNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtSMAGastosRetirosNoviembre.TabIndex = 1379;
			this.txtSMAGastosRetirosNoviembre.Visible = false;
			// 
			// txtSMAGastosRetirosOctubre
			// 
			appearance89.BackColorDisabled = System.Drawing.Color.White;
			appearance89.BackColorDisabled2 = System.Drawing.Color.White;
			appearance89.BorderColor = System.Drawing.Color.Black;
			appearance89.BorderColor3DBase = System.Drawing.Color.Black;
			appearance89.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			appearance89.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.txtSMAGastosRetirosOctubre.Appearance = appearance89;
			this.txtSMAGastosRetirosOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSMAGastosRetirosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSMAGastosRetirosOctubre.Enabled = false;
			this.txtSMAGastosRetirosOctubre.FormatString = "#,##0.00";
			this.txtSMAGastosRetirosOctubre.Location = new System.Drawing.Point(810, 504);
			this.txtSMAGastosRetirosOctubre.Name = "txtSMAGastosRetirosOctubre";
			this.txtSMAGastosRetirosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSMAGastosRetirosOctubre.PromptChar = ' ';
			this.txtSMAGastosRetirosOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtSMAGastosRetirosOctubre.TabIndex = 1378;
			this.txtSMAGastosRetirosOctubre.Visible = false;
			// 
			// txtSMAGastosRetirosSeptiembre
			// 
			appearance90.BackColorDisabled = System.Drawing.Color.White;
			appearance90.BackColorDisabled2 = System.Drawing.Color.White;
			appearance90.BorderColor = System.Drawing.Color.Black;
			appearance90.BorderColor3DBase = System.Drawing.Color.Black;
			appearance90.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			appearance90.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.txtSMAGastosRetirosSeptiembre.Appearance = appearance90;
			this.txtSMAGastosRetirosSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSMAGastosRetirosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSMAGastosRetirosSeptiembre.Enabled = false;
			this.txtSMAGastosRetirosSeptiembre.FormatString = "#,##0.00";
			this.txtSMAGastosRetirosSeptiembre.Location = new System.Drawing.Point(722, 504);
			this.txtSMAGastosRetirosSeptiembre.Name = "txtSMAGastosRetirosSeptiembre";
			this.txtSMAGastosRetirosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSMAGastosRetirosSeptiembre.PromptChar = ' ';
			this.txtSMAGastosRetirosSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtSMAGastosRetirosSeptiembre.TabIndex = 1377;
			this.txtSMAGastosRetirosSeptiembre.Visible = false;
			// 
			// btnActualizar
			// 
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnActualizar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.Location = new System.Drawing.Point(1112, 5);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(30, 23);
			this.btnActualizar.TabIndex = 1376;
			this.btnActualizar.Visible = false;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// txtSMAGastosRetirosAgosto
			// 
			appearance91.BackColorDisabled = System.Drawing.Color.White;
			appearance91.BackColorDisabled2 = System.Drawing.Color.White;
			appearance91.BorderColor = System.Drawing.Color.Black;
			appearance91.BorderColor3DBase = System.Drawing.Color.Black;
			appearance91.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			appearance91.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.txtSMAGastosRetirosAgosto.Appearance = appearance91;
			this.txtSMAGastosRetirosAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSMAGastosRetirosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSMAGastosRetirosAgosto.Enabled = false;
			this.txtSMAGastosRetirosAgosto.FormatString = "#,##0.00";
			this.txtSMAGastosRetirosAgosto.Location = new System.Drawing.Point(634, 504);
			this.txtSMAGastosRetirosAgosto.Name = "txtSMAGastosRetirosAgosto";
			this.txtSMAGastosRetirosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSMAGastosRetirosAgosto.PromptChar = ' ';
			this.txtSMAGastosRetirosAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtSMAGastosRetirosAgosto.TabIndex = 1375;
			this.txtSMAGastosRetirosAgosto.Visible = false;
			// 
			// txtSMAGastosRetirosJulio
			// 
			appearance92.BackColorDisabled = System.Drawing.Color.White;
			appearance92.BackColorDisabled2 = System.Drawing.Color.White;
			appearance92.BorderColor = System.Drawing.Color.Black;
			appearance92.BorderColor3DBase = System.Drawing.Color.Black;
			appearance92.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			appearance92.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.txtSMAGastosRetirosJulio.Appearance = appearance92;
			this.txtSMAGastosRetirosJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSMAGastosRetirosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSMAGastosRetirosJulio.Enabled = false;
			this.txtSMAGastosRetirosJulio.FormatString = "#,##0.00";
			this.txtSMAGastosRetirosJulio.Location = new System.Drawing.Point(546, 504);
			this.txtSMAGastosRetirosJulio.Name = "txtSMAGastosRetirosJulio";
			this.txtSMAGastosRetirosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSMAGastosRetirosJulio.PromptChar = ' ';
			this.txtSMAGastosRetirosJulio.Size = new System.Drawing.Size(80, 19);
			this.txtSMAGastosRetirosJulio.TabIndex = 1374;
			this.txtSMAGastosRetirosJulio.Visible = false;
			// 
			// txtSMAGastosRetirosJunio
			// 
			appearance93.BackColorDisabled = System.Drawing.Color.White;
			appearance93.BackColorDisabled2 = System.Drawing.Color.White;
			appearance93.BorderColor = System.Drawing.Color.Black;
			appearance93.BorderColor3DBase = System.Drawing.Color.Black;
			appearance93.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			appearance93.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.txtSMAGastosRetirosJunio.Appearance = appearance93;
			this.txtSMAGastosRetirosJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSMAGastosRetirosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSMAGastosRetirosJunio.Enabled = false;
			this.txtSMAGastosRetirosJunio.FormatString = "#,##0.00";
			this.txtSMAGastosRetirosJunio.Location = new System.Drawing.Point(458, 504);
			this.txtSMAGastosRetirosJunio.Name = "txtSMAGastosRetirosJunio";
			this.txtSMAGastosRetirosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSMAGastosRetirosJunio.PromptChar = ' ';
			this.txtSMAGastosRetirosJunio.Size = new System.Drawing.Size(80, 19);
			this.txtSMAGastosRetirosJunio.TabIndex = 1373;
			this.txtSMAGastosRetirosJunio.Visible = false;
			// 
			// txtSMAGastosRetirosMayo
			// 
			appearance94.BackColorDisabled = System.Drawing.Color.White;
			appearance94.BackColorDisabled2 = System.Drawing.Color.White;
			appearance94.BorderColor = System.Drawing.Color.Black;
			appearance94.BorderColor3DBase = System.Drawing.Color.Black;
			appearance94.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			appearance94.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.txtSMAGastosRetirosMayo.Appearance = appearance94;
			this.txtSMAGastosRetirosMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSMAGastosRetirosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSMAGastosRetirosMayo.Enabled = false;
			this.txtSMAGastosRetirosMayo.FormatString = "#,##0.00";
			this.txtSMAGastosRetirosMayo.Location = new System.Drawing.Point(370, 504);
			this.txtSMAGastosRetirosMayo.Name = "txtSMAGastosRetirosMayo";
			this.txtSMAGastosRetirosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSMAGastosRetirosMayo.PromptChar = ' ';
			this.txtSMAGastosRetirosMayo.Size = new System.Drawing.Size(80, 19);
			this.txtSMAGastosRetirosMayo.TabIndex = 1372;
			this.txtSMAGastosRetirosMayo.Visible = false;
			// 
			// txtSMAGastosRetirosAbril
			// 
			appearance95.BackColorDisabled = System.Drawing.Color.White;
			appearance95.BackColorDisabled2 = System.Drawing.Color.White;
			appearance95.BorderColor = System.Drawing.Color.Black;
			appearance95.BorderColor3DBase = System.Drawing.Color.Black;
			appearance95.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			appearance95.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.txtSMAGastosRetirosAbril.Appearance = appearance95;
			this.txtSMAGastosRetirosAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSMAGastosRetirosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSMAGastosRetirosAbril.Enabled = false;
			this.txtSMAGastosRetirosAbril.FormatString = "#,##0.00";
			this.txtSMAGastosRetirosAbril.Location = new System.Drawing.Point(282, 504);
			this.txtSMAGastosRetirosAbril.Name = "txtSMAGastosRetirosAbril";
			this.txtSMAGastosRetirosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSMAGastosRetirosAbril.PromptChar = ' ';
			this.txtSMAGastosRetirosAbril.Size = new System.Drawing.Size(80, 19);
			this.txtSMAGastosRetirosAbril.TabIndex = 1371;
			this.txtSMAGastosRetirosAbril.Visible = false;
			// 
			// txtSMAGastosRetirosMarzo
			// 
			appearance96.BackColorDisabled = System.Drawing.Color.White;
			appearance96.BackColorDisabled2 = System.Drawing.Color.White;
			appearance96.BorderColor = System.Drawing.Color.Black;
			appearance96.BorderColor3DBase = System.Drawing.Color.Black;
			appearance96.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			appearance96.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.txtSMAGastosRetirosMarzo.Appearance = appearance96;
			this.txtSMAGastosRetirosMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSMAGastosRetirosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSMAGastosRetirosMarzo.Enabled = false;
			this.txtSMAGastosRetirosMarzo.FormatString = "#,##0.00";
			this.txtSMAGastosRetirosMarzo.Location = new System.Drawing.Point(194, 504);
			this.txtSMAGastosRetirosMarzo.Name = "txtSMAGastosRetirosMarzo";
			this.txtSMAGastosRetirosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSMAGastosRetirosMarzo.PromptChar = ' ';
			this.txtSMAGastosRetirosMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtSMAGastosRetirosMarzo.TabIndex = 1370;
			this.txtSMAGastosRetirosMarzo.Visible = false;
			// 
			// txtSMAGastosRetirosFebrero
			// 
			appearance97.BackColorDisabled = System.Drawing.Color.White;
			appearance97.BackColorDisabled2 = System.Drawing.Color.White;
			appearance97.BorderColor = System.Drawing.Color.Black;
			appearance97.BorderColor3DBase = System.Drawing.Color.Black;
			appearance97.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			appearance97.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.txtSMAGastosRetirosFebrero.Appearance = appearance97;
			this.txtSMAGastosRetirosFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtSMAGastosRetirosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSMAGastosRetirosFebrero.Enabled = false;
			this.txtSMAGastosRetirosFebrero.FormatString = "#,##0.00";
			this.txtSMAGastosRetirosFebrero.Location = new System.Drawing.Point(106, 504);
			this.txtSMAGastosRetirosFebrero.Name = "txtSMAGastosRetirosFebrero";
			this.txtSMAGastosRetirosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSMAGastosRetirosFebrero.PromptChar = ' ';
			this.txtSMAGastosRetirosFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtSMAGastosRetirosFebrero.TabIndex = 1369;
			this.txtSMAGastosRetirosFebrero.Visible = false;
			// 
			// label26
			// 
			this.label26.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label26.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label26.ForeColor = System.Drawing.Color.Black;
			this.label26.Location = new System.Drawing.Point(0, 376);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(1218, 1);
			this.label26.TabIndex = 1368;
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label31.ForeColor = System.Drawing.Color.Firebrick;
			this.label31.Location = new System.Drawing.Point(10, 640);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(34, 16);
			this.label31.TabIndex = 1367;
			this.label31.Text = "= U N";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label30.ForeColor = System.Drawing.Color.Firebrick;
			this.label30.Location = new System.Drawing.Point(13, 616);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(28, 16);
			this.label30.TabIndex = 1366;
			this.label30.Text = "(-) R";
			this.label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label29.ForeColor = System.Drawing.Color.Black;
			this.label29.Location = new System.Drawing.Point(16, 544);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(23, 16);
			this.label29.TabIndex = 1365;
			this.label29.Text = "= G";
			this.label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label28.ForeColor = System.Drawing.Color.Firebrick;
			this.label28.Location = new System.Drawing.Point(16, 384);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(18, 16);
			this.label28.TabIndex = 1364;
			this.label28.Text = "= I";
			this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label23.ForeColor = System.Drawing.Color.Firebrick;
			this.label23.Location = new System.Drawing.Point(10, 304);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(23, 16);
			this.label23.TabIndex = 1269;
			this.label23.Text = "+ C";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label22.ForeColor = System.Drawing.Color.Firebrick;
			this.label22.Location = new System.Drawing.Point(10, 168);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(23, 16);
			this.label22.TabIndex = 1268;
			this.label22.Text = "+ C";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label21.ForeColor = System.Drawing.Color.Firebrick;
			this.label21.Location = new System.Drawing.Point(11, 280);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(21, 16);
			this.label21.TabIndex = 1267;
			this.label21.Text = "+ T";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label20.ForeColor = System.Drawing.Color.Firebrick;
			this.label20.Location = new System.Drawing.Point(11, 144);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(21, 16);
			this.label20.TabIndex = 1266;
			this.label20.Text = "+ T";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label19.ForeColor = System.Drawing.Color.Firebrick;
			this.label19.Location = new System.Drawing.Point(10, 257);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(22, 16);
			this.label19.TabIndex = 1265;
			this.label19.Text = "+ E";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.ForeColor = System.Drawing.Color.Firebrick;
			this.label18.Location = new System.Drawing.Point(10, 120);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(22, 16);
			this.label18.TabIndex = 1264;
			this.label18.Text = "+ E";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.ForeColor = System.Drawing.Color.Firebrick;
			this.label17.Location = new System.Drawing.Point(10, 232);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(23, 16);
			this.label17.TabIndex = 1263;
			this.label17.Text = "+ D";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(10, 96);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(23, 16);
			this.label5.TabIndex = 1262;
			this.label5.Text = "+ D";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotalGastos
			// 
			appearance98.BackColorDisabled = System.Drawing.Color.White;
			appearance98.BackColorDisabled2 = System.Drawing.Color.White;
			appearance98.BorderColor = System.Drawing.Color.Black;
			appearance98.BorderColor3DBase = System.Drawing.Color.Black;
			appearance98.FontData.BoldAsString = "True";
			appearance98.ForeColor = System.Drawing.Color.Black;
			appearance98.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastos.Appearance = appearance98;
			this.txtTotalGastos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastos.Enabled = false;
			this.txtTotalGastos.FormatString = "#,##0.00";
			this.txtTotalGastos.Location = new System.Drawing.Point(1114, 544);
			this.txtTotalGastos.Name = "txtTotalGastos";
			this.txtTotalGastos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastos.PromptChar = ' ';
			this.txtTotalGastos.Size = new System.Drawing.Size(100, 19);
			this.txtTotalGastos.TabIndex = 1363;
			// 
			// txtTotalGastosNoviembre
			// 
			appearance99.BackColorDisabled = System.Drawing.Color.White;
			appearance99.BackColorDisabled2 = System.Drawing.Color.White;
			appearance99.BorderColor = System.Drawing.Color.Black;
			appearance99.BorderColor3DBase = System.Drawing.Color.Black;
			appearance99.ForeColor = System.Drawing.Color.Black;
			appearance99.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosNoviembre.Appearance = appearance99;
			this.txtTotalGastosNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosNoviembre.Enabled = false;
			this.txtTotalGastosNoviembre.FormatString = "#,##0.00";
			this.txtTotalGastosNoviembre.Location = new System.Drawing.Point(938, 544);
			this.txtTotalGastosNoviembre.Name = "txtTotalGastosNoviembre";
			this.txtTotalGastosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosNoviembre.PromptChar = ' ';
			this.txtTotalGastosNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosNoviembre.TabIndex = 1362;
			// 
			// txtTotalGastosDiciembre
			// 
			appearance100.BackColorDisabled = System.Drawing.Color.White;
			appearance100.BackColorDisabled2 = System.Drawing.Color.White;
			appearance100.BorderColor = System.Drawing.Color.Black;
			appearance100.BorderColor3DBase = System.Drawing.Color.Black;
			appearance100.ForeColor = System.Drawing.Color.Black;
			appearance100.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosDiciembre.Appearance = appearance100;
			this.txtTotalGastosDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosDiciembre.Enabled = false;
			this.txtTotalGastosDiciembre.FormatString = "#,##0.00";
			this.txtTotalGastosDiciembre.Location = new System.Drawing.Point(1026, 544);
			this.txtTotalGastosDiciembre.Name = "txtTotalGastosDiciembre";
			this.txtTotalGastosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosDiciembre.PromptChar = ' ';
			this.txtTotalGastosDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosDiciembre.TabIndex = 1361;
			// 
			// txtTotalGastosSeptiembre
			// 
			appearance101.BackColorDisabled = System.Drawing.Color.White;
			appearance101.BackColorDisabled2 = System.Drawing.Color.White;
			appearance101.BorderColor = System.Drawing.Color.Black;
			appearance101.BorderColor3DBase = System.Drawing.Color.Black;
			appearance101.ForeColor = System.Drawing.Color.Black;
			appearance101.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosSeptiembre.Appearance = appearance101;
			this.txtTotalGastosSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosSeptiembre.Enabled = false;
			this.txtTotalGastosSeptiembre.FormatString = "#,##0.00";
			this.txtTotalGastosSeptiembre.Location = new System.Drawing.Point(762, 544);
			this.txtTotalGastosSeptiembre.Name = "txtTotalGastosSeptiembre";
			this.txtTotalGastosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosSeptiembre.PromptChar = ' ';
			this.txtTotalGastosSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosSeptiembre.TabIndex = 1360;
			// 
			// txtTotalGastosOctubre
			// 
			appearance102.BackColorDisabled = System.Drawing.Color.White;
			appearance102.BackColorDisabled2 = System.Drawing.Color.White;
			appearance102.BorderColor = System.Drawing.Color.Black;
			appearance102.BorderColor3DBase = System.Drawing.Color.Black;
			appearance102.ForeColor = System.Drawing.Color.Black;
			appearance102.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosOctubre.Appearance = appearance102;
			this.txtTotalGastosOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosOctubre.Enabled = false;
			this.txtTotalGastosOctubre.FormatString = "#,##0.00";
			this.txtTotalGastosOctubre.Location = new System.Drawing.Point(850, 544);
			this.txtTotalGastosOctubre.Name = "txtTotalGastosOctubre";
			this.txtTotalGastosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosOctubre.PromptChar = ' ';
			this.txtTotalGastosOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosOctubre.TabIndex = 1359;
			// 
			// txtTotalGastosJulio
			// 
			appearance103.BackColorDisabled = System.Drawing.Color.White;
			appearance103.BackColorDisabled2 = System.Drawing.Color.White;
			appearance103.BorderColor = System.Drawing.Color.Black;
			appearance103.BorderColor3DBase = System.Drawing.Color.Black;
			appearance103.ForeColor = System.Drawing.Color.Black;
			appearance103.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosJulio.Appearance = appearance103;
			this.txtTotalGastosJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosJulio.Enabled = false;
			this.txtTotalGastosJulio.FormatString = "#,##0.00";
			this.txtTotalGastosJulio.Location = new System.Drawing.Point(586, 544);
			this.txtTotalGastosJulio.Name = "txtTotalGastosJulio";
			this.txtTotalGastosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosJulio.PromptChar = ' ';
			this.txtTotalGastosJulio.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosJulio.TabIndex = 1358;
			// 
			// txtTotalGastosAgosto
			// 
			appearance104.BackColorDisabled = System.Drawing.Color.White;
			appearance104.BackColorDisabled2 = System.Drawing.Color.White;
			appearance104.BorderColor = System.Drawing.Color.Black;
			appearance104.BorderColor3DBase = System.Drawing.Color.Black;
			appearance104.ForeColor = System.Drawing.Color.Black;
			appearance104.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosAgosto.Appearance = appearance104;
			this.txtTotalGastosAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosAgosto.Enabled = false;
			this.txtTotalGastosAgosto.FormatString = "#,##0.00";
			this.txtTotalGastosAgosto.Location = new System.Drawing.Point(674, 544);
			this.txtTotalGastosAgosto.Name = "txtTotalGastosAgosto";
			this.txtTotalGastosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosAgosto.PromptChar = ' ';
			this.txtTotalGastosAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosAgosto.TabIndex = 1357;
			// 
			// txtTotalGastosMayo
			// 
			appearance105.BackColorDisabled = System.Drawing.Color.White;
			appearance105.BackColorDisabled2 = System.Drawing.Color.White;
			appearance105.BorderColor = System.Drawing.Color.Black;
			appearance105.BorderColor3DBase = System.Drawing.Color.Black;
			appearance105.ForeColor = System.Drawing.Color.Black;
			appearance105.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosMayo.Appearance = appearance105;
			this.txtTotalGastosMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosMayo.Enabled = false;
			this.txtTotalGastosMayo.FormatString = "#,##0.00";
			this.txtTotalGastosMayo.Location = new System.Drawing.Point(410, 544);
			this.txtTotalGastosMayo.Name = "txtTotalGastosMayo";
			this.txtTotalGastosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosMayo.PromptChar = ' ';
			this.txtTotalGastosMayo.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosMayo.TabIndex = 1356;
			// 
			// txtTotalGastosJunio
			// 
			appearance106.BackColorDisabled = System.Drawing.Color.White;
			appearance106.BackColorDisabled2 = System.Drawing.Color.White;
			appearance106.BorderColor = System.Drawing.Color.Black;
			appearance106.BorderColor3DBase = System.Drawing.Color.Black;
			appearance106.ForeColor = System.Drawing.Color.Black;
			appearance106.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosJunio.Appearance = appearance106;
			this.txtTotalGastosJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosJunio.Enabled = false;
			this.txtTotalGastosJunio.FormatString = "#,##0.00";
			this.txtTotalGastosJunio.Location = new System.Drawing.Point(498, 544);
			this.txtTotalGastosJunio.Name = "txtTotalGastosJunio";
			this.txtTotalGastosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosJunio.PromptChar = ' ';
			this.txtTotalGastosJunio.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosJunio.TabIndex = 1355;
			// 
			// txtTotalGastosMarzo
			// 
			appearance107.BackColorDisabled = System.Drawing.Color.White;
			appearance107.BackColorDisabled2 = System.Drawing.Color.White;
			appearance107.BorderColor = System.Drawing.Color.Black;
			appearance107.BorderColor3DBase = System.Drawing.Color.Black;
			appearance107.ForeColor = System.Drawing.Color.Black;
			appearance107.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosMarzo.Appearance = appearance107;
			this.txtTotalGastosMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosMarzo.Enabled = false;
			this.txtTotalGastosMarzo.FormatString = "#,##0.00";
			this.txtTotalGastosMarzo.Location = new System.Drawing.Point(234, 544);
			this.txtTotalGastosMarzo.Name = "txtTotalGastosMarzo";
			this.txtTotalGastosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosMarzo.PromptChar = ' ';
			this.txtTotalGastosMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosMarzo.TabIndex = 1354;
			// 
			// txtTotalGastosAbril
			// 
			appearance108.BackColorDisabled = System.Drawing.Color.White;
			appearance108.BackColorDisabled2 = System.Drawing.Color.White;
			appearance108.BorderColor = System.Drawing.Color.Black;
			appearance108.BorderColor3DBase = System.Drawing.Color.Black;
			appearance108.ForeColor = System.Drawing.Color.Black;
			appearance108.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosAbril.Appearance = appearance108;
			this.txtTotalGastosAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosAbril.Enabled = false;
			this.txtTotalGastosAbril.FormatString = "#,##0.00";
			this.txtTotalGastosAbril.Location = new System.Drawing.Point(322, 544);
			this.txtTotalGastosAbril.Name = "txtTotalGastosAbril";
			this.txtTotalGastosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosAbril.PromptChar = ' ';
			this.txtTotalGastosAbril.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosAbril.TabIndex = 1353;
			// 
			// txtTotalGastosFebrero
			// 
			appearance109.BackColorDisabled = System.Drawing.Color.White;
			appearance109.BackColorDisabled2 = System.Drawing.Color.White;
			appearance109.BorderColor = System.Drawing.Color.Black;
			appearance109.BorderColor3DBase = System.Drawing.Color.Black;
			appearance109.ForeColor = System.Drawing.Color.Black;
			appearance109.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosFebrero.Appearance = appearance109;
			this.txtTotalGastosFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosFebrero.Enabled = false;
			this.txtTotalGastosFebrero.FormatString = "#,##0.00";
			this.txtTotalGastosFebrero.Location = new System.Drawing.Point(146, 544);
			this.txtTotalGastosFebrero.Name = "txtTotalGastosFebrero";
			this.txtTotalGastosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosFebrero.PromptChar = ' ';
			this.txtTotalGastosFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosFebrero.TabIndex = 1352;
			// 
			// txtTotalGastosEnero
			// 
			appearance110.BackColorDisabled = System.Drawing.Color.White;
			appearance110.BackColorDisabled2 = System.Drawing.Color.White;
			appearance110.BorderColor = System.Drawing.Color.Black;
			appearance110.BorderColor3DBase = System.Drawing.Color.Black;
			appearance110.ForeColor = System.Drawing.Color.Black;
			appearance110.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastosEnero.Appearance = appearance110;
			this.txtTotalGastosEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalGastosEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastosEnero.Enabled = false;
			this.txtTotalGastosEnero.FormatString = "#,##0.00";
			this.txtTotalGastosEnero.Location = new System.Drawing.Point(58, 544);
			this.txtTotalGastosEnero.Name = "txtTotalGastosEnero";
			this.txtTotalGastosEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastosEnero.PromptChar = ' ';
			this.txtTotalGastosEnero.Size = new System.Drawing.Size(80, 19);
			this.txtTotalGastosEnero.TabIndex = 1351;
			// 
			// txtResultadoTotal
			// 
			appearance111.BackColorDisabled = System.Drawing.Color.White;
			appearance111.BackColorDisabled2 = System.Drawing.Color.White;
			appearance111.BorderColor = System.Drawing.Color.Black;
			appearance111.BorderColor3DBase = System.Drawing.Color.Black;
			appearance111.FontData.BoldAsString = "True";
			appearance111.ForeColor = System.Drawing.Color.Firebrick;
			appearance111.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoTotal.Appearance = appearance111;
			this.txtResultadoTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoTotal.Enabled = false;
			this.txtResultadoTotal.FormatString = "#,##0.00";
			this.txtResultadoTotal.Location = new System.Drawing.Point(1114, 640);
			this.txtResultadoTotal.Name = "txtResultadoTotal";
			this.txtResultadoTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoTotal.PromptChar = ' ';
			this.txtResultadoTotal.Size = new System.Drawing.Size(100, 19);
			this.txtResultadoTotal.TabIndex = 1350;
			// 
			// txtResultadoNoviembre
			// 
			appearance112.BackColorDisabled = System.Drawing.Color.White;
			appearance112.BackColorDisabled2 = System.Drawing.Color.White;
			appearance112.BorderColor = System.Drawing.Color.Black;
			appearance112.BorderColor3DBase = System.Drawing.Color.Black;
			appearance112.ForeColor = System.Drawing.Color.Firebrick;
			appearance112.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoNoviembre.Appearance = appearance112;
			this.txtResultadoNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoNoviembre.Enabled = false;
			this.txtResultadoNoviembre.FormatString = "#,##0.00";
			this.txtResultadoNoviembre.Location = new System.Drawing.Point(938, 640);
			this.txtResultadoNoviembre.Name = "txtResultadoNoviembre";
			this.txtResultadoNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoNoviembre.PromptChar = ' ';
			this.txtResultadoNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoNoviembre.TabIndex = 1349;
			// 
			// txtResultadoDiciembre
			// 
			appearance113.BackColorDisabled = System.Drawing.Color.White;
			appearance113.BackColorDisabled2 = System.Drawing.Color.White;
			appearance113.BorderColor = System.Drawing.Color.Black;
			appearance113.BorderColor3DBase = System.Drawing.Color.Black;
			appearance113.ForeColor = System.Drawing.Color.Firebrick;
			appearance113.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoDiciembre.Appearance = appearance113;
			this.txtResultadoDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoDiciembre.Enabled = false;
			this.txtResultadoDiciembre.FormatString = "#,##0.00";
			this.txtResultadoDiciembre.Location = new System.Drawing.Point(1026, 640);
			this.txtResultadoDiciembre.Name = "txtResultadoDiciembre";
			this.txtResultadoDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoDiciembre.PromptChar = ' ';
			this.txtResultadoDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoDiciembre.TabIndex = 1348;
			// 
			// txtResultadoSeptiembre
			// 
			appearance114.BackColorDisabled = System.Drawing.Color.White;
			appearance114.BackColorDisabled2 = System.Drawing.Color.White;
			appearance114.BorderColor = System.Drawing.Color.Black;
			appearance114.BorderColor3DBase = System.Drawing.Color.Black;
			appearance114.ForeColor = System.Drawing.Color.Firebrick;
			appearance114.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoSeptiembre.Appearance = appearance114;
			this.txtResultadoSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoSeptiembre.Enabled = false;
			this.txtResultadoSeptiembre.FormatString = "#,##0.00";
			this.txtResultadoSeptiembre.Location = new System.Drawing.Point(762, 640);
			this.txtResultadoSeptiembre.Name = "txtResultadoSeptiembre";
			this.txtResultadoSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoSeptiembre.PromptChar = ' ';
			this.txtResultadoSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoSeptiembre.TabIndex = 1347;
			// 
			// txtResultadoOctubre
			// 
			appearance115.BackColorDisabled = System.Drawing.Color.White;
			appearance115.BackColorDisabled2 = System.Drawing.Color.White;
			appearance115.BorderColor = System.Drawing.Color.Black;
			appearance115.BorderColor3DBase = System.Drawing.Color.Black;
			appearance115.ForeColor = System.Drawing.Color.Firebrick;
			appearance115.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoOctubre.Appearance = appearance115;
			this.txtResultadoOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoOctubre.Enabled = false;
			this.txtResultadoOctubre.FormatString = "#,##0.00";
			this.txtResultadoOctubre.Location = new System.Drawing.Point(850, 640);
			this.txtResultadoOctubre.Name = "txtResultadoOctubre";
			this.txtResultadoOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoOctubre.PromptChar = ' ';
			this.txtResultadoOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoOctubre.TabIndex = 1346;
			// 
			// txtResultadoJulio
			// 
			appearance116.BackColorDisabled = System.Drawing.Color.White;
			appearance116.BackColorDisabled2 = System.Drawing.Color.White;
			appearance116.BorderColor = System.Drawing.Color.Black;
			appearance116.BorderColor3DBase = System.Drawing.Color.Black;
			appearance116.ForeColor = System.Drawing.Color.Firebrick;
			appearance116.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoJulio.Appearance = appearance116;
			this.txtResultadoJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoJulio.Enabled = false;
			this.txtResultadoJulio.FormatString = "#,##0.00";
			this.txtResultadoJulio.Location = new System.Drawing.Point(586, 640);
			this.txtResultadoJulio.Name = "txtResultadoJulio";
			this.txtResultadoJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoJulio.PromptChar = ' ';
			this.txtResultadoJulio.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoJulio.TabIndex = 1345;
			// 
			// txtResultadoAgosto
			// 
			appearance117.BackColorDisabled = System.Drawing.Color.White;
			appearance117.BackColorDisabled2 = System.Drawing.Color.White;
			appearance117.BorderColor = System.Drawing.Color.Black;
			appearance117.BorderColor3DBase = System.Drawing.Color.Black;
			appearance117.ForeColor = System.Drawing.Color.Firebrick;
			appearance117.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoAgosto.Appearance = appearance117;
			this.txtResultadoAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoAgosto.Enabled = false;
			this.txtResultadoAgosto.FormatString = "#,##0.00";
			this.txtResultadoAgosto.Location = new System.Drawing.Point(674, 640);
			this.txtResultadoAgosto.Name = "txtResultadoAgosto";
			this.txtResultadoAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoAgosto.PromptChar = ' ';
			this.txtResultadoAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoAgosto.TabIndex = 1344;
			// 
			// txtResultadoMayo
			// 
			appearance118.BackColorDisabled = System.Drawing.Color.White;
			appearance118.BackColorDisabled2 = System.Drawing.Color.White;
			appearance118.BorderColor = System.Drawing.Color.Black;
			appearance118.BorderColor3DBase = System.Drawing.Color.Black;
			appearance118.ForeColor = System.Drawing.Color.Firebrick;
			appearance118.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoMayo.Appearance = appearance118;
			this.txtResultadoMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoMayo.Enabled = false;
			this.txtResultadoMayo.FormatString = "#,##0.00";
			this.txtResultadoMayo.Location = new System.Drawing.Point(410, 640);
			this.txtResultadoMayo.Name = "txtResultadoMayo";
			this.txtResultadoMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoMayo.PromptChar = ' ';
			this.txtResultadoMayo.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoMayo.TabIndex = 1343;
			// 
			// txtResultadoJunio
			// 
			appearance119.BackColorDisabled = System.Drawing.Color.White;
			appearance119.BackColorDisabled2 = System.Drawing.Color.White;
			appearance119.BorderColor = System.Drawing.Color.Black;
			appearance119.BorderColor3DBase = System.Drawing.Color.Black;
			appearance119.ForeColor = System.Drawing.Color.Firebrick;
			appearance119.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoJunio.Appearance = appearance119;
			this.txtResultadoJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoJunio.Enabled = false;
			this.txtResultadoJunio.FormatString = "#,##0.00";
			this.txtResultadoJunio.Location = new System.Drawing.Point(498, 640);
			this.txtResultadoJunio.Name = "txtResultadoJunio";
			this.txtResultadoJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoJunio.PromptChar = ' ';
			this.txtResultadoJunio.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoJunio.TabIndex = 1342;
			// 
			// txtResultadoMarzo
			// 
			appearance120.BackColorDisabled = System.Drawing.Color.White;
			appearance120.BackColorDisabled2 = System.Drawing.Color.White;
			appearance120.BorderColor = System.Drawing.Color.Black;
			appearance120.BorderColor3DBase = System.Drawing.Color.Black;
			appearance120.ForeColor = System.Drawing.Color.Firebrick;
			appearance120.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoMarzo.Appearance = appearance120;
			this.txtResultadoMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoMarzo.Enabled = false;
			this.txtResultadoMarzo.FormatString = "#,##0.00";
			this.txtResultadoMarzo.Location = new System.Drawing.Point(234, 640);
			this.txtResultadoMarzo.Name = "txtResultadoMarzo";
			this.txtResultadoMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoMarzo.PromptChar = ' ';
			this.txtResultadoMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoMarzo.TabIndex = 1341;
			// 
			// txtResultadoAbril
			// 
			appearance121.BackColorDisabled = System.Drawing.Color.White;
			appearance121.BackColorDisabled2 = System.Drawing.Color.White;
			appearance121.BorderColor = System.Drawing.Color.Black;
			appearance121.BorderColor3DBase = System.Drawing.Color.Black;
			appearance121.ForeColor = System.Drawing.Color.Firebrick;
			appearance121.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoAbril.Appearance = appearance121;
			this.txtResultadoAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoAbril.Enabled = false;
			this.txtResultadoAbril.FormatString = "#,##0.00";
			this.txtResultadoAbril.Location = new System.Drawing.Point(322, 640);
			this.txtResultadoAbril.Name = "txtResultadoAbril";
			this.txtResultadoAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoAbril.PromptChar = ' ';
			this.txtResultadoAbril.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoAbril.TabIndex = 1340;
			// 
			// txtResultadoFebrero
			// 
			appearance122.BackColorDisabled = System.Drawing.Color.White;
			appearance122.BackColorDisabled2 = System.Drawing.Color.White;
			appearance122.BorderColor = System.Drawing.Color.Black;
			appearance122.BorderColor3DBase = System.Drawing.Color.Black;
			appearance122.ForeColor = System.Drawing.Color.Firebrick;
			appearance122.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoFebrero.Appearance = appearance122;
			this.txtResultadoFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoFebrero.Enabled = false;
			this.txtResultadoFebrero.FormatString = "#,##0.00";
			this.txtResultadoFebrero.Location = new System.Drawing.Point(146, 640);
			this.txtResultadoFebrero.Name = "txtResultadoFebrero";
			this.txtResultadoFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoFebrero.PromptChar = ' ';
			this.txtResultadoFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoFebrero.TabIndex = 1339;
			// 
			// txtResultadoEnero
			// 
			appearance123.BackColorDisabled = System.Drawing.Color.White;
			appearance123.BackColorDisabled2 = System.Drawing.Color.White;
			appearance123.BorderColor = System.Drawing.Color.Black;
			appearance123.BorderColor3DBase = System.Drawing.Color.Black;
			appearance123.ForeColor = System.Drawing.Color.Firebrick;
			appearance123.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtResultadoEnero.Appearance = appearance123;
			this.txtResultadoEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtResultadoEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultadoEnero.Enabled = false;
			this.txtResultadoEnero.FormatString = "#,##0.00";
			this.txtResultadoEnero.Location = new System.Drawing.Point(58, 640);
			this.txtResultadoEnero.Name = "txtResultadoEnero";
			this.txtResultadoEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtResultadoEnero.PromptChar = ' ';
			this.txtResultadoEnero.Size = new System.Drawing.Size(80, 19);
			this.txtResultadoEnero.TabIndex = 1338;
			// 
			// txtGastosRetiroTotal
			// 
			appearance124.BackColorDisabled = System.Drawing.Color.White;
			appearance124.BackColorDisabled2 = System.Drawing.Color.White;
			appearance124.BorderColor = System.Drawing.Color.Black;
			appearance124.BorderColor3DBase = System.Drawing.Color.Black;
			appearance124.FontData.BoldAsString = "True";
			appearance124.ForeColor = System.Drawing.Color.Black;
			appearance124.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroTotal.Appearance = appearance124;
			this.txtGastosRetiroTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroTotal.Enabled = false;
			this.txtGastosRetiroTotal.FormatString = "#,##0.00";
			this.txtGastosRetiroTotal.Location = new System.Drawing.Point(1114, 616);
			this.txtGastosRetiroTotal.Name = "txtGastosRetiroTotal";
			this.txtGastosRetiroTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroTotal.PromptChar = ' ';
			this.txtGastosRetiroTotal.ReadOnly = true;
			this.txtGastosRetiroTotal.Size = new System.Drawing.Size(100, 19);
			this.txtGastosRetiroTotal.TabIndex = 1337;
			this.txtGastosRetiroTotal.DoubleClick += new System.EventHandler(this.txtGastosRetiroTotal_DoubleClick);
			// 
			// txtGastosRetiroNoviembre
			// 
			appearance125.BackColorDisabled = System.Drawing.Color.White;
			appearance125.BackColorDisabled2 = System.Drawing.Color.White;
			appearance125.BorderColor = System.Drawing.Color.Black;
			appearance125.BorderColor3DBase = System.Drawing.Color.Black;
			appearance125.ForeColor = System.Drawing.Color.Black;
			appearance125.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroNoviembre.Appearance = appearance125;
			this.txtGastosRetiroNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroNoviembre.FormatString = "#,##0.00";
			this.txtGastosRetiroNoviembre.Location = new System.Drawing.Point(938, 616);
			this.txtGastosRetiroNoviembre.Name = "txtGastosRetiroNoviembre";
			this.txtGastosRetiroNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroNoviembre.PromptChar = ' ';
			this.txtGastosRetiroNoviembre.ReadOnly = true;
			this.txtGastosRetiroNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroNoviembre.TabIndex = 1336;
			this.txtGastosRetiroNoviembre.DoubleClick += new System.EventHandler(this.txtGastosRetiroNoviembre_DoubleClick);
			// 
			// txtGastosRetiroDiciembre
			// 
			appearance126.BackColorDisabled = System.Drawing.Color.White;
			appearance126.BackColorDisabled2 = System.Drawing.Color.White;
			appearance126.BorderColor = System.Drawing.Color.Black;
			appearance126.BorderColor3DBase = System.Drawing.Color.Black;
			appearance126.ForeColor = System.Drawing.Color.Black;
			appearance126.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroDiciembre.Appearance = appearance126;
			this.txtGastosRetiroDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroDiciembre.FormatString = "#,##0.00";
			this.txtGastosRetiroDiciembre.Location = new System.Drawing.Point(1026, 616);
			this.txtGastosRetiroDiciembre.Name = "txtGastosRetiroDiciembre";
			this.txtGastosRetiroDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroDiciembre.PromptChar = ' ';
			this.txtGastosRetiroDiciembre.ReadOnly = true;
			this.txtGastosRetiroDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroDiciembre.TabIndex = 1335;
			this.txtGastosRetiroDiciembre.DoubleClick += new System.EventHandler(this.txtGastosRetiroDiciembre_DoubleClick);
			// 
			// txtGastosRetiroSeptiembre
			// 
			appearance127.BackColorDisabled = System.Drawing.Color.White;
			appearance127.BackColorDisabled2 = System.Drawing.Color.White;
			appearance127.BorderColor = System.Drawing.Color.Black;
			appearance127.BorderColor3DBase = System.Drawing.Color.Black;
			appearance127.ForeColor = System.Drawing.Color.Black;
			appearance127.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroSeptiembre.Appearance = appearance127;
			this.txtGastosRetiroSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroSeptiembre.FormatString = "#,##0.00";
			this.txtGastosRetiroSeptiembre.Location = new System.Drawing.Point(762, 616);
			this.txtGastosRetiroSeptiembre.Name = "txtGastosRetiroSeptiembre";
			this.txtGastosRetiroSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroSeptiembre.PromptChar = ' ';
			this.txtGastosRetiroSeptiembre.ReadOnly = true;
			this.txtGastosRetiroSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroSeptiembre.TabIndex = 1334;
			this.txtGastosRetiroSeptiembre.DoubleClick += new System.EventHandler(this.txtGastosRetiroSeptiembre_DoubleClick);
			// 
			// txtGastosRetiroOctubre
			// 
			appearance128.BackColorDisabled = System.Drawing.Color.White;
			appearance128.BackColorDisabled2 = System.Drawing.Color.White;
			appearance128.BorderColor = System.Drawing.Color.Black;
			appearance128.BorderColor3DBase = System.Drawing.Color.Black;
			appearance128.ForeColor = System.Drawing.Color.Black;
			appearance128.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroOctubre.Appearance = appearance128;
			this.txtGastosRetiroOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroOctubre.FormatString = "#,##0.00";
			this.txtGastosRetiroOctubre.Location = new System.Drawing.Point(850, 616);
			this.txtGastosRetiroOctubre.Name = "txtGastosRetiroOctubre";
			this.txtGastosRetiroOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroOctubre.PromptChar = ' ';
			this.txtGastosRetiroOctubre.ReadOnly = true;
			this.txtGastosRetiroOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroOctubre.TabIndex = 1333;
			this.txtGastosRetiroOctubre.DoubleClick += new System.EventHandler(this.txtGastosRetiroOctubre_DoubleClick);
			// 
			// txtGastosRetiroJulio
			// 
			appearance129.BackColorDisabled = System.Drawing.Color.White;
			appearance129.BackColorDisabled2 = System.Drawing.Color.White;
			appearance129.BorderColor = System.Drawing.Color.Black;
			appearance129.BorderColor3DBase = System.Drawing.Color.Black;
			appearance129.ForeColor = System.Drawing.Color.Black;
			appearance129.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroJulio.Appearance = appearance129;
			this.txtGastosRetiroJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroJulio.FormatString = "#,##0.00";
			this.txtGastosRetiroJulio.Location = new System.Drawing.Point(586, 616);
			this.txtGastosRetiroJulio.Name = "txtGastosRetiroJulio";
			this.txtGastosRetiroJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroJulio.PromptChar = ' ';
			this.txtGastosRetiroJulio.ReadOnly = true;
			this.txtGastosRetiroJulio.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroJulio.TabIndex = 1332;
			this.txtGastosRetiroJulio.DoubleClick += new System.EventHandler(this.txtGastosRetiroJulio_DoubleClick);
			// 
			// txtGastosRetiroAgosto
			// 
			appearance130.BackColorDisabled = System.Drawing.Color.White;
			appearance130.BackColorDisabled2 = System.Drawing.Color.White;
			appearance130.BorderColor = System.Drawing.Color.Black;
			appearance130.BorderColor3DBase = System.Drawing.Color.Black;
			appearance130.ForeColor = System.Drawing.Color.Black;
			appearance130.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroAgosto.Appearance = appearance130;
			this.txtGastosRetiroAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroAgosto.FormatString = "#,##0.00";
			this.txtGastosRetiroAgosto.Location = new System.Drawing.Point(674, 616);
			this.txtGastosRetiroAgosto.Name = "txtGastosRetiroAgosto";
			this.txtGastosRetiroAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroAgosto.PromptChar = ' ';
			this.txtGastosRetiroAgosto.ReadOnly = true;
			this.txtGastosRetiroAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroAgosto.TabIndex = 1331;
			this.txtGastosRetiroAgosto.DoubleClick += new System.EventHandler(this.txtGastosRetiroAgosto_DoubleClick);
			// 
			// txtGastosRetiroMayo
			// 
			appearance131.BackColorDisabled = System.Drawing.Color.White;
			appearance131.BackColorDisabled2 = System.Drawing.Color.White;
			appearance131.BorderColor = System.Drawing.Color.Black;
			appearance131.BorderColor3DBase = System.Drawing.Color.Black;
			appearance131.ForeColor = System.Drawing.Color.Black;
			appearance131.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroMayo.Appearance = appearance131;
			this.txtGastosRetiroMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroMayo.FormatString = "#,##0.00";
			this.txtGastosRetiroMayo.Location = new System.Drawing.Point(410, 616);
			this.txtGastosRetiroMayo.Name = "txtGastosRetiroMayo";
			this.txtGastosRetiroMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroMayo.PromptChar = ' ';
			this.txtGastosRetiroMayo.ReadOnly = true;
			this.txtGastosRetiroMayo.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroMayo.TabIndex = 1330;
			this.txtGastosRetiroMayo.DoubleClick += new System.EventHandler(this.txtGastosRetiroMayo_DoubleClick);
			// 
			// txtGastosRetiroJunio
			// 
			appearance132.BackColorDisabled = System.Drawing.Color.White;
			appearance132.BackColorDisabled2 = System.Drawing.Color.White;
			appearance132.BorderColor = System.Drawing.Color.Black;
			appearance132.BorderColor3DBase = System.Drawing.Color.Black;
			appearance132.ForeColor = System.Drawing.Color.Black;
			appearance132.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroJunio.Appearance = appearance132;
			this.txtGastosRetiroJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroJunio.FormatString = "#,##0.00";
			this.txtGastosRetiroJunio.Location = new System.Drawing.Point(498, 616);
			this.txtGastosRetiroJunio.Name = "txtGastosRetiroJunio";
			this.txtGastosRetiroJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroJunio.PromptChar = ' ';
			this.txtGastosRetiroJunio.ReadOnly = true;
			this.txtGastosRetiroJunio.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroJunio.TabIndex = 1329;
			this.txtGastosRetiroJunio.DoubleClick += new System.EventHandler(this.txtGastosRetiroJunio_DoubleClick);
			// 
			// txtGastosRetiroMarzo
			// 
			appearance133.BackColorDisabled = System.Drawing.Color.White;
			appearance133.BackColorDisabled2 = System.Drawing.Color.White;
			appearance133.BorderColor = System.Drawing.Color.Black;
			appearance133.BorderColor3DBase = System.Drawing.Color.Black;
			appearance133.ForeColor = System.Drawing.Color.Black;
			appearance133.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroMarzo.Appearance = appearance133;
			this.txtGastosRetiroMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroMarzo.FormatString = "#,##0.00";
			this.txtGastosRetiroMarzo.Location = new System.Drawing.Point(234, 616);
			this.txtGastosRetiroMarzo.Name = "txtGastosRetiroMarzo";
			this.txtGastosRetiroMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroMarzo.PromptChar = ' ';
			this.txtGastosRetiroMarzo.ReadOnly = true;
			this.txtGastosRetiroMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroMarzo.TabIndex = 1328;
			this.txtGastosRetiroMarzo.DoubleClick += new System.EventHandler(this.txtGastosRetiroMarzo_DoubleClick);
			// 
			// txtGastosRetiroAbril
			// 
			appearance134.BackColorDisabled = System.Drawing.Color.White;
			appearance134.BackColorDisabled2 = System.Drawing.Color.White;
			appearance134.BorderColor = System.Drawing.Color.Black;
			appearance134.BorderColor3DBase = System.Drawing.Color.Black;
			appearance134.ForeColor = System.Drawing.Color.Black;
			appearance134.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroAbril.Appearance = appearance134;
			this.txtGastosRetiroAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroAbril.FormatString = "#,##0.00";
			this.txtGastosRetiroAbril.Location = new System.Drawing.Point(322, 616);
			this.txtGastosRetiroAbril.Name = "txtGastosRetiroAbril";
			this.txtGastosRetiroAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroAbril.PromptChar = ' ';
			this.txtGastosRetiroAbril.ReadOnly = true;
			this.txtGastosRetiroAbril.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroAbril.TabIndex = 1327;
			this.txtGastosRetiroAbril.DoubleClick += new System.EventHandler(this.txtGastosRetiroAbril_DoubleClick);
			// 
			// txtGastosRetiroFebrero
			// 
			appearance135.BackColorDisabled = System.Drawing.Color.White;
			appearance135.BackColorDisabled2 = System.Drawing.Color.White;
			appearance135.BorderColor = System.Drawing.Color.Black;
			appearance135.BorderColor3DBase = System.Drawing.Color.Black;
			appearance135.ForeColor = System.Drawing.Color.Black;
			appearance135.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroFebrero.Appearance = appearance135;
			this.txtGastosRetiroFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroFebrero.FormatString = "#,##0.00";
			this.txtGastosRetiroFebrero.Location = new System.Drawing.Point(146, 616);
			this.txtGastosRetiroFebrero.Name = "txtGastosRetiroFebrero";
			this.txtGastosRetiroFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroFebrero.PromptChar = ' ';
			this.txtGastosRetiroFebrero.ReadOnly = true;
			this.txtGastosRetiroFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroFebrero.TabIndex = 1326;
			this.txtGastosRetiroFebrero.DoubleClick += new System.EventHandler(this.txtGastosRetiroFebrero_DoubleClick);
			// 
			// txtGastosRetiroEnero
			// 
			appearance136.BackColorDisabled = System.Drawing.Color.White;
			appearance136.BackColorDisabled2 = System.Drawing.Color.White;
			appearance136.BorderColor = System.Drawing.Color.Black;
			appearance136.BorderColor3DBase = System.Drawing.Color.Black;
			appearance136.ForeColor = System.Drawing.Color.Black;
			appearance136.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosRetiroEnero.Appearance = appearance136;
			this.txtGastosRetiroEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtGastosRetiroEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosRetiroEnero.Enabled = false;
			this.txtGastosRetiroEnero.FormatString = "#,##0.00";
			this.txtGastosRetiroEnero.Location = new System.Drawing.Point(58, 616);
			this.txtGastosRetiroEnero.Name = "txtGastosRetiroEnero";
			this.txtGastosRetiroEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosRetiroEnero.PromptChar = ' ';
			this.txtGastosRetiroEnero.ReadOnly = true;
			this.txtGastosRetiroEnero.Size = new System.Drawing.Size(80, 19);
			this.txtGastosRetiroEnero.TabIndex = 1325;
			this.txtGastosRetiroEnero.DoubleClick += new System.EventHandler(this.txtGastosRetiroEnero_DoubleClick);
			// 
			// txtTotalIngresos
			// 
			appearance137.BackColorDisabled = System.Drawing.Color.White;
			appearance137.BackColorDisabled2 = System.Drawing.Color.White;
			appearance137.BorderColor = System.Drawing.Color.Black;
			appearance137.BorderColor3DBase = System.Drawing.Color.Black;
			appearance137.FontData.BoldAsString = "True";
			appearance137.ForeColor = System.Drawing.Color.Firebrick;
			appearance137.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresos.Appearance = appearance137;
			this.txtTotalIngresos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresos.Enabled = false;
			this.txtTotalIngresos.FormatString = "#,##0.00";
			this.txtTotalIngresos.Location = new System.Drawing.Point(1114, 376);
			this.txtTotalIngresos.Name = "txtTotalIngresos";
			this.txtTotalIngresos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresos.PromptChar = ' ';
			this.txtTotalIngresos.Size = new System.Drawing.Size(100, 19);
			this.txtTotalIngresos.TabIndex = 1324;
			// 
			// txtTotalIngresosNoviembre
			// 
			appearance138.BackColorDisabled = System.Drawing.Color.White;
			appearance138.BackColorDisabled2 = System.Drawing.Color.White;
			appearance138.BorderColor = System.Drawing.Color.Black;
			appearance138.BorderColor3DBase = System.Drawing.Color.Black;
			appearance138.ForeColor = System.Drawing.Color.Firebrick;
			appearance138.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosNoviembre.Appearance = appearance138;
			this.txtTotalIngresosNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosNoviembre.Enabled = false;
			this.txtTotalIngresosNoviembre.FormatString = "#,##0.00";
			this.txtTotalIngresosNoviembre.Location = new System.Drawing.Point(938, 376);
			this.txtTotalIngresosNoviembre.Name = "txtTotalIngresosNoviembre";
			this.txtTotalIngresosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosNoviembre.PromptChar = ' ';
			this.txtTotalIngresosNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosNoviembre.TabIndex = 1323;
			// 
			// txtTotalIngresosDiciembre
			// 
			appearance139.BackColorDisabled = System.Drawing.Color.White;
			appearance139.BackColorDisabled2 = System.Drawing.Color.White;
			appearance139.BorderColor = System.Drawing.Color.Black;
			appearance139.BorderColor3DBase = System.Drawing.Color.Black;
			appearance139.ForeColor = System.Drawing.Color.Firebrick;
			appearance139.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosDiciembre.Appearance = appearance139;
			this.txtTotalIngresosDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosDiciembre.Enabled = false;
			this.txtTotalIngresosDiciembre.FormatString = "#,##0.00";
			this.txtTotalIngresosDiciembre.Location = new System.Drawing.Point(1026, 376);
			this.txtTotalIngresosDiciembre.Name = "txtTotalIngresosDiciembre";
			this.txtTotalIngresosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosDiciembre.PromptChar = ' ';
			this.txtTotalIngresosDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosDiciembre.TabIndex = 1322;
			// 
			// txtTotalIngresosSeptiembre
			// 
			appearance140.BackColorDisabled = System.Drawing.Color.White;
			appearance140.BackColorDisabled2 = System.Drawing.Color.White;
			appearance140.BorderColor = System.Drawing.Color.Black;
			appearance140.BorderColor3DBase = System.Drawing.Color.Black;
			appearance140.ForeColor = System.Drawing.Color.Firebrick;
			appearance140.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosSeptiembre.Appearance = appearance140;
			this.txtTotalIngresosSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosSeptiembre.Enabled = false;
			this.txtTotalIngresosSeptiembre.FormatString = "#,##0.00";
			this.txtTotalIngresosSeptiembre.Location = new System.Drawing.Point(762, 376);
			this.txtTotalIngresosSeptiembre.Name = "txtTotalIngresosSeptiembre";
			this.txtTotalIngresosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosSeptiembre.PromptChar = ' ';
			this.txtTotalIngresosSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosSeptiembre.TabIndex = 1321;
			// 
			// txtTotalIngresosOctubre
			// 
			appearance141.BackColorDisabled = System.Drawing.Color.White;
			appearance141.BackColorDisabled2 = System.Drawing.Color.White;
			appearance141.BorderColor = System.Drawing.Color.Black;
			appearance141.BorderColor3DBase = System.Drawing.Color.Black;
			appearance141.ForeColor = System.Drawing.Color.Firebrick;
			appearance141.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosOctubre.Appearance = appearance141;
			this.txtTotalIngresosOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosOctubre.Enabled = false;
			this.txtTotalIngresosOctubre.FormatString = "#,##0.00";
			this.txtTotalIngresosOctubre.Location = new System.Drawing.Point(850, 376);
			this.txtTotalIngresosOctubre.Name = "txtTotalIngresosOctubre";
			this.txtTotalIngresosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosOctubre.PromptChar = ' ';
			this.txtTotalIngresosOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosOctubre.TabIndex = 1320;
			// 
			// txtTotalIngresosJulio
			// 
			appearance142.BackColorDisabled = System.Drawing.Color.White;
			appearance142.BackColorDisabled2 = System.Drawing.Color.White;
			appearance142.BorderColor = System.Drawing.Color.Black;
			appearance142.BorderColor3DBase = System.Drawing.Color.Black;
			appearance142.ForeColor = System.Drawing.Color.Firebrick;
			appearance142.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosJulio.Appearance = appearance142;
			this.txtTotalIngresosJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosJulio.Enabled = false;
			this.txtTotalIngresosJulio.FormatString = "#,##0.00";
			this.txtTotalIngresosJulio.Location = new System.Drawing.Point(586, 376);
			this.txtTotalIngresosJulio.Name = "txtTotalIngresosJulio";
			this.txtTotalIngresosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosJulio.PromptChar = ' ';
			this.txtTotalIngresosJulio.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosJulio.TabIndex = 1319;
			// 
			// txtTotalIngresosAgosto
			// 
			appearance143.BackColorDisabled = System.Drawing.Color.White;
			appearance143.BackColorDisabled2 = System.Drawing.Color.White;
			appearance143.BorderColor = System.Drawing.Color.Black;
			appearance143.BorderColor3DBase = System.Drawing.Color.Black;
			appearance143.ForeColor = System.Drawing.Color.Firebrick;
			appearance143.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosAgosto.Appearance = appearance143;
			this.txtTotalIngresosAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosAgosto.Enabled = false;
			this.txtTotalIngresosAgosto.FormatString = "#,##0.00";
			this.txtTotalIngresosAgosto.Location = new System.Drawing.Point(674, 376);
			this.txtTotalIngresosAgosto.Name = "txtTotalIngresosAgosto";
			this.txtTotalIngresosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosAgosto.PromptChar = ' ';
			this.txtTotalIngresosAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosAgosto.TabIndex = 1318;
			// 
			// txtTotalIngresosMayo
			// 
			appearance144.BackColorDisabled = System.Drawing.Color.White;
			appearance144.BackColorDisabled2 = System.Drawing.Color.White;
			appearance144.BorderColor = System.Drawing.Color.Black;
			appearance144.BorderColor3DBase = System.Drawing.Color.Black;
			appearance144.ForeColor = System.Drawing.Color.Firebrick;
			appearance144.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosMayo.Appearance = appearance144;
			this.txtTotalIngresosMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosMayo.Enabled = false;
			this.txtTotalIngresosMayo.FormatString = "#,##0.00";
			this.txtTotalIngresosMayo.Location = new System.Drawing.Point(410, 376);
			this.txtTotalIngresosMayo.Name = "txtTotalIngresosMayo";
			this.txtTotalIngresosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosMayo.PromptChar = ' ';
			this.txtTotalIngresosMayo.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosMayo.TabIndex = 1317;
			// 
			// txtTotalIngresosJunio
			// 
			appearance145.BackColorDisabled = System.Drawing.Color.White;
			appearance145.BackColorDisabled2 = System.Drawing.Color.White;
			appearance145.BorderColor = System.Drawing.Color.Black;
			appearance145.BorderColor3DBase = System.Drawing.Color.Black;
			appearance145.ForeColor = System.Drawing.Color.Firebrick;
			appearance145.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosJunio.Appearance = appearance145;
			this.txtTotalIngresosJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosJunio.Enabled = false;
			this.txtTotalIngresosJunio.FormatString = "#,##0.00";
			this.txtTotalIngresosJunio.Location = new System.Drawing.Point(498, 376);
			this.txtTotalIngresosJunio.Name = "txtTotalIngresosJunio";
			this.txtTotalIngresosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosJunio.PromptChar = ' ';
			this.txtTotalIngresosJunio.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosJunio.TabIndex = 1316;
			// 
			// txtTotalIngresosMarzo
			// 
			appearance146.BackColorDisabled = System.Drawing.Color.White;
			appearance146.BackColorDisabled2 = System.Drawing.Color.White;
			appearance146.BorderColor = System.Drawing.Color.Black;
			appearance146.BorderColor3DBase = System.Drawing.Color.Black;
			appearance146.ForeColor = System.Drawing.Color.Firebrick;
			appearance146.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosMarzo.Appearance = appearance146;
			this.txtTotalIngresosMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosMarzo.Enabled = false;
			this.txtTotalIngresosMarzo.FormatString = "#,##0.00";
			this.txtTotalIngresosMarzo.Location = new System.Drawing.Point(234, 376);
			this.txtTotalIngresosMarzo.Name = "txtTotalIngresosMarzo";
			this.txtTotalIngresosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosMarzo.PromptChar = ' ';
			this.txtTotalIngresosMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosMarzo.TabIndex = 1315;
			// 
			// txtTotalIngresosAbril
			// 
			appearance147.BackColorDisabled = System.Drawing.Color.White;
			appearance147.BackColorDisabled2 = System.Drawing.Color.White;
			appearance147.BorderColor = System.Drawing.Color.Black;
			appearance147.BorderColor3DBase = System.Drawing.Color.Black;
			appearance147.ForeColor = System.Drawing.Color.Firebrick;
			appearance147.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosAbril.Appearance = appearance147;
			this.txtTotalIngresosAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosAbril.Enabled = false;
			this.txtTotalIngresosAbril.FormatString = "#,##0.00";
			this.txtTotalIngresosAbril.Location = new System.Drawing.Point(322, 376);
			this.txtTotalIngresosAbril.Name = "txtTotalIngresosAbril";
			this.txtTotalIngresosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosAbril.PromptChar = ' ';
			this.txtTotalIngresosAbril.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosAbril.TabIndex = 1314;
			// 
			// txtTotalIngresosFebrero
			// 
			appearance148.BackColorDisabled = System.Drawing.Color.White;
			appearance148.BackColorDisabled2 = System.Drawing.Color.White;
			appearance148.BorderColor = System.Drawing.Color.Black;
			appearance148.BorderColor3DBase = System.Drawing.Color.Black;
			appearance148.ForeColor = System.Drawing.Color.Firebrick;
			appearance148.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosFebrero.Appearance = appearance148;
			this.txtTotalIngresosFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosFebrero.Enabled = false;
			this.txtTotalIngresosFebrero.FormatString = "#,##0.00";
			this.txtTotalIngresosFebrero.Location = new System.Drawing.Point(146, 376);
			this.txtTotalIngresosFebrero.Name = "txtTotalIngresosFebrero";
			this.txtTotalIngresosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosFebrero.PromptChar = ' ';
			this.txtTotalIngresosFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosFebrero.TabIndex = 1313;
			// 
			// txtTotalIngresosEnero
			// 
			appearance149.BackColorDisabled = System.Drawing.Color.White;
			appearance149.BackColorDisabled2 = System.Drawing.Color.White;
			appearance149.BorderColor = System.Drawing.Color.Black;
			appearance149.BorderColor3DBase = System.Drawing.Color.Black;
			appearance149.ForeColor = System.Drawing.Color.Firebrick;
			appearance149.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotalIngresosEnero.Appearance = appearance149;
			this.txtTotalIngresosEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtTotalIngresosEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresosEnero.Enabled = false;
			this.txtTotalIngresosEnero.FormatString = "#,##0.00";
			this.txtTotalIngresosEnero.Location = new System.Drawing.Point(58, 376);
			this.txtTotalIngresosEnero.Name = "txtTotalIngresosEnero";
			this.txtTotalIngresosEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresosEnero.PromptChar = ' ';
			this.txtTotalIngresosEnero.Size = new System.Drawing.Size(80, 19);
			this.txtTotalIngresosEnero.TabIndex = 1312;
			// 
			// txtCobranzaTotal
			// 
			appearance150.BackColorDisabled = System.Drawing.Color.White;
			appearance150.BackColorDisabled2 = System.Drawing.Color.White;
			appearance150.BorderColor = System.Drawing.Color.Black;
			appearance150.BorderColor3DBase = System.Drawing.Color.Black;
			appearance150.FontData.BoldAsString = "True";
			appearance150.ForeColor = System.Drawing.Color.Firebrick;
			appearance150.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtCobranzaTotal.Appearance = appearance150;
			this.txtCobranzaTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaTotal.Enabled = false;
			this.txtCobranzaTotal.FormatString = "#,##0.00";
			this.txtCobranzaTotal.Location = new System.Drawing.Point(1114, 352);
			this.txtCobranzaTotal.Name = "txtCobranzaTotal";
			this.txtCobranzaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaTotal.PromptChar = ' ';
			this.txtCobranzaTotal.Size = new System.Drawing.Size(100, 19);
			this.txtCobranzaTotal.TabIndex = 1311;
			// 
			// txtCobranzaNoviembre
			// 
			appearance151.BackColorDisabled = System.Drawing.Color.White;
			appearance151.BackColorDisabled2 = System.Drawing.Color.White;
			appearance151.BorderColor = System.Drawing.Color.Black;
			appearance151.BorderColor3DBase = System.Drawing.Color.Black;
			appearance151.ForeColor = System.Drawing.Color.Black;
			appearance151.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaNoviembre.Appearance = appearance151;
			this.txtCobranzaNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaNoviembre.Enabled = false;
			this.txtCobranzaNoviembre.FormatString = "#,##0.00";
			this.txtCobranzaNoviembre.Location = new System.Drawing.Point(938, 352);
			this.txtCobranzaNoviembre.Name = "txtCobranzaNoviembre";
			this.txtCobranzaNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaNoviembre.PromptChar = ' ';
			this.txtCobranzaNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaNoviembre.TabIndex = 1310;
			// 
			// txtCobranzaDiciembre
			// 
			appearance152.BackColorDisabled = System.Drawing.Color.White;
			appearance152.BackColorDisabled2 = System.Drawing.Color.White;
			appearance152.BorderColor = System.Drawing.Color.Black;
			appearance152.BorderColor3DBase = System.Drawing.Color.Black;
			appearance152.ForeColor = System.Drawing.Color.Black;
			appearance152.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaDiciembre.Appearance = appearance152;
			this.txtCobranzaDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaDiciembre.Enabled = false;
			this.txtCobranzaDiciembre.FormatString = "#,##0.00";
			this.txtCobranzaDiciembre.Location = new System.Drawing.Point(1026, 352);
			this.txtCobranzaDiciembre.Name = "txtCobranzaDiciembre";
			this.txtCobranzaDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaDiciembre.PromptChar = ' ';
			this.txtCobranzaDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaDiciembre.TabIndex = 1309;
			// 
			// txtCobranzaSeptiembre
			// 
			appearance153.BackColorDisabled = System.Drawing.Color.White;
			appearance153.BackColorDisabled2 = System.Drawing.Color.White;
			appearance153.BorderColor = System.Drawing.Color.Black;
			appearance153.BorderColor3DBase = System.Drawing.Color.Black;
			appearance153.ForeColor = System.Drawing.Color.Black;
			appearance153.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaSeptiembre.Appearance = appearance153;
			this.txtCobranzaSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaSeptiembre.Enabled = false;
			this.txtCobranzaSeptiembre.FormatString = "#,##0.00";
			this.txtCobranzaSeptiembre.Location = new System.Drawing.Point(762, 352);
			this.txtCobranzaSeptiembre.Name = "txtCobranzaSeptiembre";
			this.txtCobranzaSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaSeptiembre.PromptChar = ' ';
			this.txtCobranzaSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaSeptiembre.TabIndex = 1308;
			// 
			// txtCobranzaOctubre
			// 
			appearance154.BackColorDisabled = System.Drawing.Color.White;
			appearance154.BackColorDisabled2 = System.Drawing.Color.White;
			appearance154.BorderColor = System.Drawing.Color.Black;
			appearance154.BorderColor3DBase = System.Drawing.Color.Black;
			appearance154.ForeColor = System.Drawing.Color.Black;
			appearance154.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaOctubre.Appearance = appearance154;
			this.txtCobranzaOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaOctubre.Enabled = false;
			this.txtCobranzaOctubre.FormatString = "#,##0.00";
			this.txtCobranzaOctubre.Location = new System.Drawing.Point(850, 352);
			this.txtCobranzaOctubre.Name = "txtCobranzaOctubre";
			this.txtCobranzaOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaOctubre.PromptChar = ' ';
			this.txtCobranzaOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaOctubre.TabIndex = 1307;
			// 
			// txtCobranzaJulio
			// 
			appearance155.BackColorDisabled = System.Drawing.Color.White;
			appearance155.BackColorDisabled2 = System.Drawing.Color.White;
			appearance155.BorderColor = System.Drawing.Color.Black;
			appearance155.BorderColor3DBase = System.Drawing.Color.Black;
			appearance155.ForeColor = System.Drawing.Color.Black;
			appearance155.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaJulio.Appearance = appearance155;
			this.txtCobranzaJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaJulio.Enabled = false;
			this.txtCobranzaJulio.FormatString = "#,##0.00";
			this.txtCobranzaJulio.Location = new System.Drawing.Point(586, 352);
			this.txtCobranzaJulio.Name = "txtCobranzaJulio";
			this.txtCobranzaJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaJulio.PromptChar = ' ';
			this.txtCobranzaJulio.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaJulio.TabIndex = 1306;
			// 
			// txtCobranzaAgosto
			// 
			appearance156.BackColorDisabled = System.Drawing.Color.White;
			appearance156.BackColorDisabled2 = System.Drawing.Color.White;
			appearance156.BorderColor = System.Drawing.Color.Black;
			appearance156.BorderColor3DBase = System.Drawing.Color.Black;
			appearance156.ForeColor = System.Drawing.Color.Black;
			appearance156.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaAgosto.Appearance = appearance156;
			this.txtCobranzaAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaAgosto.Enabled = false;
			this.txtCobranzaAgosto.FormatString = "#,##0.00";
			this.txtCobranzaAgosto.Location = new System.Drawing.Point(674, 352);
			this.txtCobranzaAgosto.Name = "txtCobranzaAgosto";
			this.txtCobranzaAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaAgosto.PromptChar = ' ';
			this.txtCobranzaAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaAgosto.TabIndex = 1305;
			// 
			// txtCobranzaMayo
			// 
			appearance157.BackColorDisabled = System.Drawing.Color.White;
			appearance157.BackColorDisabled2 = System.Drawing.Color.White;
			appearance157.BorderColor = System.Drawing.Color.Black;
			appearance157.BorderColor3DBase = System.Drawing.Color.Black;
			appearance157.ForeColor = System.Drawing.Color.Black;
			appearance157.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaMayo.Appearance = appearance157;
			this.txtCobranzaMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaMayo.Enabled = false;
			this.txtCobranzaMayo.FormatString = "#,##0.00";
			this.txtCobranzaMayo.Location = new System.Drawing.Point(410, 352);
			this.txtCobranzaMayo.Name = "txtCobranzaMayo";
			this.txtCobranzaMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaMayo.PromptChar = ' ';
			this.txtCobranzaMayo.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaMayo.TabIndex = 1304;
			// 
			// txtCobranzaJunio
			// 
			appearance158.BackColorDisabled = System.Drawing.Color.White;
			appearance158.BackColorDisabled2 = System.Drawing.Color.White;
			appearance158.BorderColor = System.Drawing.Color.Black;
			appearance158.BorderColor3DBase = System.Drawing.Color.Black;
			appearance158.ForeColor = System.Drawing.Color.Black;
			appearance158.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaJunio.Appearance = appearance158;
			this.txtCobranzaJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaJunio.Enabled = false;
			this.txtCobranzaJunio.FormatString = "#,##0.00";
			this.txtCobranzaJunio.Location = new System.Drawing.Point(498, 352);
			this.txtCobranzaJunio.Name = "txtCobranzaJunio";
			this.txtCobranzaJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaJunio.PromptChar = ' ';
			this.txtCobranzaJunio.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaJunio.TabIndex = 1303;
			// 
			// txtCobranzaMarzo
			// 
			appearance159.BackColorDisabled = System.Drawing.Color.White;
			appearance159.BackColorDisabled2 = System.Drawing.Color.White;
			appearance159.BorderColor = System.Drawing.Color.Black;
			appearance159.BorderColor3DBase = System.Drawing.Color.Black;
			appearance159.ForeColor = System.Drawing.Color.Black;
			appearance159.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaMarzo.Appearance = appearance159;
			this.txtCobranzaMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaMarzo.Enabled = false;
			this.txtCobranzaMarzo.FormatString = "#,##0.00";
			this.txtCobranzaMarzo.Location = new System.Drawing.Point(234, 352);
			this.txtCobranzaMarzo.Name = "txtCobranzaMarzo";
			this.txtCobranzaMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaMarzo.PromptChar = ' ';
			this.txtCobranzaMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaMarzo.TabIndex = 1302;
			// 
			// txtCobranzaAbril
			// 
			appearance160.BackColorDisabled = System.Drawing.Color.White;
			appearance160.BackColorDisabled2 = System.Drawing.Color.White;
			appearance160.BorderColor = System.Drawing.Color.Black;
			appearance160.BorderColor3DBase = System.Drawing.Color.Black;
			appearance160.ForeColor = System.Drawing.Color.Black;
			appearance160.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaAbril.Appearance = appearance160;
			this.txtCobranzaAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaAbril.Enabled = false;
			this.txtCobranzaAbril.FormatString = "#,##0.00";
			this.txtCobranzaAbril.Location = new System.Drawing.Point(322, 352);
			this.txtCobranzaAbril.Name = "txtCobranzaAbril";
			this.txtCobranzaAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaAbril.PromptChar = ' ';
			this.txtCobranzaAbril.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaAbril.TabIndex = 1301;
			// 
			// txtCobranzaFebrero
			// 
			appearance161.BackColorDisabled = System.Drawing.Color.White;
			appearance161.BackColorDisabled2 = System.Drawing.Color.White;
			appearance161.BorderColor = System.Drawing.Color.Black;
			appearance161.BorderColor3DBase = System.Drawing.Color.Black;
			appearance161.ForeColor = System.Drawing.Color.Black;
			appearance161.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaFebrero.Appearance = appearance161;
			this.txtCobranzaFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaFebrero.Enabled = false;
			this.txtCobranzaFebrero.FormatString = "#,##0.00";
			this.txtCobranzaFebrero.Location = new System.Drawing.Point(146, 352);
			this.txtCobranzaFebrero.Name = "txtCobranzaFebrero";
			this.txtCobranzaFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaFebrero.PromptChar = ' ';
			this.txtCobranzaFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaFebrero.TabIndex = 1300;
			// 
			// txtCobranzaEnero
			// 
			appearance162.BackColorDisabled = System.Drawing.Color.White;
			appearance162.BackColorDisabled2 = System.Drawing.Color.White;
			appearance162.BorderColor = System.Drawing.Color.Black;
			appearance162.BorderColor3DBase = System.Drawing.Color.Black;
			appearance162.ForeColor = System.Drawing.Color.Black;
			appearance162.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobranzaEnero.Appearance = appearance162;
			this.txtCobranzaEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCobranzaEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobranzaEnero.Enabled = false;
			this.txtCobranzaEnero.FormatString = "#,##0.00";
			this.txtCobranzaEnero.Location = new System.Drawing.Point(58, 352);
			this.txtCobranzaEnero.Name = "txtCobranzaEnero";
			this.txtCobranzaEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobranzaEnero.PromptChar = ' ';
			this.txtCobranzaEnero.Size = new System.Drawing.Size(80, 19);
			this.txtCobranzaEnero.TabIndex = 1299;
			// 
			// txtProyeccionMargenTotal
			// 
			appearance163.BackColorDisabled = System.Drawing.Color.White;
			appearance163.BackColorDisabled2 = System.Drawing.Color.White;
			appearance163.BorderColor = System.Drawing.Color.Black;
			appearance163.BorderColor3DBase = System.Drawing.Color.Black;
			appearance163.FontData.BoldAsString = "True";
			appearance163.ForeColor = System.Drawing.Color.Firebrick;
			appearance163.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotal.Appearance = appearance163;
			this.txtProyeccionMargenTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotal.Enabled = false;
			this.txtProyeccionMargenTotal.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotal.Location = new System.Drawing.Point(1114, 328);
			this.txtProyeccionMargenTotal.Name = "txtProyeccionMargenTotal";
			this.txtProyeccionMargenTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotal.PromptChar = ' ';
			this.txtProyeccionMargenTotal.Size = new System.Drawing.Size(100, 19);
			this.txtProyeccionMargenTotal.TabIndex = 1298;
			// 
			// txtProyeccionMargenTotalNoviembre
			// 
			appearance164.BackColorDisabled = System.Drawing.Color.White;
			appearance164.BackColorDisabled2 = System.Drawing.Color.White;
			appearance164.BorderColor = System.Drawing.Color.Black;
			appearance164.BorderColor3DBase = System.Drawing.Color.Black;
			appearance164.ForeColor = System.Drawing.Color.Firebrick;
			appearance164.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalNoviembre.Appearance = appearance164;
			this.txtProyeccionMargenTotalNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalNoviembre.Enabled = false;
			this.txtProyeccionMargenTotalNoviembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalNoviembre.Location = new System.Drawing.Point(938, 328);
			this.txtProyeccionMargenTotalNoviembre.Name = "txtProyeccionMargenTotalNoviembre";
			this.txtProyeccionMargenTotalNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalNoviembre.PromptChar = ' ';
			this.txtProyeccionMargenTotalNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalNoviembre.TabIndex = 1297;
			// 
			// txtProyeccionMargenTotalDiciembre
			// 
			appearance165.BackColorDisabled = System.Drawing.Color.White;
			appearance165.BackColorDisabled2 = System.Drawing.Color.White;
			appearance165.BorderColor = System.Drawing.Color.Black;
			appearance165.BorderColor3DBase = System.Drawing.Color.Black;
			appearance165.ForeColor = System.Drawing.Color.Firebrick;
			appearance165.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalDiciembre.Appearance = appearance165;
			this.txtProyeccionMargenTotalDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalDiciembre.Enabled = false;
			this.txtProyeccionMargenTotalDiciembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalDiciembre.Location = new System.Drawing.Point(1026, 328);
			this.txtProyeccionMargenTotalDiciembre.Name = "txtProyeccionMargenTotalDiciembre";
			this.txtProyeccionMargenTotalDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalDiciembre.PromptChar = ' ';
			this.txtProyeccionMargenTotalDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalDiciembre.TabIndex = 1296;
			// 
			// txtProyeccionMargenTotalSeptiembre
			// 
			appearance166.BackColorDisabled = System.Drawing.Color.White;
			appearance166.BackColorDisabled2 = System.Drawing.Color.White;
			appearance166.BorderColor = System.Drawing.Color.Black;
			appearance166.BorderColor3DBase = System.Drawing.Color.Black;
			appearance166.ForeColor = System.Drawing.Color.Firebrick;
			appearance166.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalSeptiembre.Appearance = appearance166;
			this.txtProyeccionMargenTotalSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalSeptiembre.Enabled = false;
			this.txtProyeccionMargenTotalSeptiembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalSeptiembre.Location = new System.Drawing.Point(762, 328);
			this.txtProyeccionMargenTotalSeptiembre.Name = "txtProyeccionMargenTotalSeptiembre";
			this.txtProyeccionMargenTotalSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalSeptiembre.PromptChar = ' ';
			this.txtProyeccionMargenTotalSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalSeptiembre.TabIndex = 1295;
			// 
			// txtProyeccionMargenTotalOctubre
			// 
			appearance167.BackColorDisabled = System.Drawing.Color.White;
			appearance167.BackColorDisabled2 = System.Drawing.Color.White;
			appearance167.BorderColor = System.Drawing.Color.Black;
			appearance167.BorderColor3DBase = System.Drawing.Color.Black;
			appearance167.ForeColor = System.Drawing.Color.Firebrick;
			appearance167.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalOctubre.Appearance = appearance167;
			this.txtProyeccionMargenTotalOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalOctubre.Enabled = false;
			this.txtProyeccionMargenTotalOctubre.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalOctubre.Location = new System.Drawing.Point(850, 328);
			this.txtProyeccionMargenTotalOctubre.Name = "txtProyeccionMargenTotalOctubre";
			this.txtProyeccionMargenTotalOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalOctubre.PromptChar = ' ';
			this.txtProyeccionMargenTotalOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalOctubre.TabIndex = 1294;
			// 
			// txtProyeccionMargenTotalJulio
			// 
			appearance168.BackColorDisabled = System.Drawing.Color.White;
			appearance168.BackColorDisabled2 = System.Drawing.Color.White;
			appearance168.BorderColor = System.Drawing.Color.Black;
			appearance168.BorderColor3DBase = System.Drawing.Color.Black;
			appearance168.ForeColor = System.Drawing.Color.Firebrick;
			appearance168.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalJulio.Appearance = appearance168;
			this.txtProyeccionMargenTotalJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalJulio.Enabled = false;
			this.txtProyeccionMargenTotalJulio.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalJulio.Location = new System.Drawing.Point(586, 328);
			this.txtProyeccionMargenTotalJulio.Name = "txtProyeccionMargenTotalJulio";
			this.txtProyeccionMargenTotalJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalJulio.PromptChar = ' ';
			this.txtProyeccionMargenTotalJulio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalJulio.TabIndex = 1293;
			// 
			// txtProyeccionMargenTotalAgosto
			// 
			appearance169.BackColorDisabled = System.Drawing.Color.White;
			appearance169.BackColorDisabled2 = System.Drawing.Color.White;
			appearance169.BorderColor = System.Drawing.Color.Black;
			appearance169.BorderColor3DBase = System.Drawing.Color.Black;
			appearance169.ForeColor = System.Drawing.Color.Firebrick;
			appearance169.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalAgosto.Appearance = appearance169;
			this.txtProyeccionMargenTotalAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalAgosto.Enabled = false;
			this.txtProyeccionMargenTotalAgosto.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalAgosto.Location = new System.Drawing.Point(674, 328);
			this.txtProyeccionMargenTotalAgosto.Name = "txtProyeccionMargenTotalAgosto";
			this.txtProyeccionMargenTotalAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalAgosto.PromptChar = ' ';
			this.txtProyeccionMargenTotalAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalAgosto.TabIndex = 1292;
			// 
			// txtProyeccionMargenTotalMayo
			// 
			appearance170.BackColorDisabled = System.Drawing.Color.White;
			appearance170.BackColorDisabled2 = System.Drawing.Color.White;
			appearance170.BorderColor = System.Drawing.Color.Black;
			appearance170.BorderColor3DBase = System.Drawing.Color.Black;
			appearance170.ForeColor = System.Drawing.Color.Firebrick;
			appearance170.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalMayo.Appearance = appearance170;
			this.txtProyeccionMargenTotalMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalMayo.Enabled = false;
			this.txtProyeccionMargenTotalMayo.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalMayo.Location = new System.Drawing.Point(410, 328);
			this.txtProyeccionMargenTotalMayo.Name = "txtProyeccionMargenTotalMayo";
			this.txtProyeccionMargenTotalMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalMayo.PromptChar = ' ';
			this.txtProyeccionMargenTotalMayo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalMayo.TabIndex = 1291;
			// 
			// txtProyeccionMargenTotalJunio
			// 
			appearance171.BackColorDisabled = System.Drawing.Color.White;
			appearance171.BackColorDisabled2 = System.Drawing.Color.White;
			appearance171.BorderColor = System.Drawing.Color.Black;
			appearance171.BorderColor3DBase = System.Drawing.Color.Black;
			appearance171.ForeColor = System.Drawing.Color.Firebrick;
			appearance171.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalJunio.Appearance = appearance171;
			this.txtProyeccionMargenTotalJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalJunio.Enabled = false;
			this.txtProyeccionMargenTotalJunio.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalJunio.Location = new System.Drawing.Point(498, 328);
			this.txtProyeccionMargenTotalJunio.Name = "txtProyeccionMargenTotalJunio";
			this.txtProyeccionMargenTotalJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalJunio.PromptChar = ' ';
			this.txtProyeccionMargenTotalJunio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalJunio.TabIndex = 1290;
			// 
			// txtProyeccionMargenTotalMarzo
			// 
			appearance172.BackColorDisabled = System.Drawing.Color.White;
			appearance172.BackColorDisabled2 = System.Drawing.Color.White;
			appearance172.BorderColor = System.Drawing.Color.Black;
			appearance172.BorderColor3DBase = System.Drawing.Color.Black;
			appearance172.ForeColor = System.Drawing.Color.Firebrick;
			appearance172.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalMarzo.Appearance = appearance172;
			this.txtProyeccionMargenTotalMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalMarzo.Enabled = false;
			this.txtProyeccionMargenTotalMarzo.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalMarzo.Location = new System.Drawing.Point(234, 328);
			this.txtProyeccionMargenTotalMarzo.Name = "txtProyeccionMargenTotalMarzo";
			this.txtProyeccionMargenTotalMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalMarzo.PromptChar = ' ';
			this.txtProyeccionMargenTotalMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalMarzo.TabIndex = 1289;
			// 
			// txtProyeccionMargenTotalAbril
			// 
			appearance173.BackColorDisabled = System.Drawing.Color.White;
			appearance173.BackColorDisabled2 = System.Drawing.Color.White;
			appearance173.BorderColor = System.Drawing.Color.Black;
			appearance173.BorderColor3DBase = System.Drawing.Color.Black;
			appearance173.ForeColor = System.Drawing.Color.Firebrick;
			appearance173.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalAbril.Appearance = appearance173;
			this.txtProyeccionMargenTotalAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalAbril.Enabled = false;
			this.txtProyeccionMargenTotalAbril.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalAbril.Location = new System.Drawing.Point(322, 328);
			this.txtProyeccionMargenTotalAbril.Name = "txtProyeccionMargenTotalAbril";
			this.txtProyeccionMargenTotalAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalAbril.PromptChar = ' ';
			this.txtProyeccionMargenTotalAbril.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalAbril.TabIndex = 1288;
			// 
			// txtProyeccionMargenTotalFebrero
			// 
			appearance174.BackColorDisabled = System.Drawing.Color.White;
			appearance174.BackColorDisabled2 = System.Drawing.Color.White;
			appearance174.BorderColor = System.Drawing.Color.Black;
			appearance174.BorderColor3DBase = System.Drawing.Color.Black;
			appearance174.ForeColor = System.Drawing.Color.Firebrick;
			appearance174.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalFebrero.Appearance = appearance174;
			this.txtProyeccionMargenTotalFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalFebrero.Enabled = false;
			this.txtProyeccionMargenTotalFebrero.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalFebrero.Location = new System.Drawing.Point(146, 328);
			this.txtProyeccionMargenTotalFebrero.Name = "txtProyeccionMargenTotalFebrero";
			this.txtProyeccionMargenTotalFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalFebrero.PromptChar = ' ';
			this.txtProyeccionMargenTotalFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalFebrero.TabIndex = 1287;
			// 
			// txtProyeccionMargenTotalEnero
			// 
			appearance175.BackColorDisabled = System.Drawing.Color.White;
			appearance175.BackColorDisabled2 = System.Drawing.Color.White;
			appearance175.BorderColor = System.Drawing.Color.Black;
			appearance175.BorderColor3DBase = System.Drawing.Color.Black;
			appearance175.ForeColor = System.Drawing.Color.Firebrick;
			appearance175.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTotalEnero.Appearance = appearance175;
			this.txtProyeccionMargenTotalEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTotalEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTotalEnero.Enabled = false;
			this.txtProyeccionMargenTotalEnero.FormatString = "#,##0.00";
			this.txtProyeccionMargenTotalEnero.Location = new System.Drawing.Point(58, 328);
			this.txtProyeccionMargenTotalEnero.Name = "txtProyeccionMargenTotalEnero";
			this.txtProyeccionMargenTotalEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTotalEnero.PromptChar = ' ';
			this.txtProyeccionMargenTotalEnero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTotalEnero.TabIndex = 1286;
			// 
			// label25
			// 
			this.label25.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label25.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label25.ForeColor = System.Drawing.Color.Black;
			this.label25.Location = new System.Drawing.Point(0, 328);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(1218, 1);
			this.label25.TabIndex = 1285;
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtProyeccionVentaTotal
			// 
			appearance176.BackColorDisabled = System.Drawing.Color.White;
			appearance176.BackColorDisabled2 = System.Drawing.Color.White;
			appearance176.BorderColor = System.Drawing.Color.Black;
			appearance176.BorderColor3DBase = System.Drawing.Color.Black;
			appearance176.FontData.BoldAsString = "True";
			appearance176.ForeColor = System.Drawing.Color.Firebrick;
			appearance176.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotal.Appearance = appearance176;
			this.txtProyeccionVentaTotal.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotal.Enabled = false;
			this.txtProyeccionVentaTotal.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotal.Location = new System.Drawing.Point(1114, 192);
			this.txtProyeccionVentaTotal.Name = "txtProyeccionVentaTotal";
			this.txtProyeccionVentaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotal.PromptChar = ' ';
			this.txtProyeccionVentaTotal.Size = new System.Drawing.Size(100, 19);
			this.txtProyeccionVentaTotal.TabIndex = 1284;
			// 
			// txtProyeccionVentaTotalNoviembre
			// 
			appearance177.BackColorDisabled = System.Drawing.Color.White;
			appearance177.BackColorDisabled2 = System.Drawing.Color.White;
			appearance177.BorderColor = System.Drawing.Color.Black;
			appearance177.BorderColor3DBase = System.Drawing.Color.Black;
			appearance177.ForeColor = System.Drawing.Color.Firebrick;
			appearance177.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalNoviembre.Appearance = appearance177;
			this.txtProyeccionVentaTotalNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalNoviembre.Enabled = false;
			this.txtProyeccionVentaTotalNoviembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalNoviembre.Location = new System.Drawing.Point(938, 192);
			this.txtProyeccionVentaTotalNoviembre.Name = "txtProyeccionVentaTotalNoviembre";
			this.txtProyeccionVentaTotalNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalNoviembre.PromptChar = ' ';
			this.txtProyeccionVentaTotalNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalNoviembre.TabIndex = 1283;
			// 
			// txtProyeccionVentaTotalDiciembre
			// 
			appearance178.BackColorDisabled = System.Drawing.Color.White;
			appearance178.BackColorDisabled2 = System.Drawing.Color.White;
			appearance178.BorderColor = System.Drawing.Color.Black;
			appearance178.BorderColor3DBase = System.Drawing.Color.Black;
			appearance178.ForeColor = System.Drawing.Color.Firebrick;
			appearance178.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalDiciembre.Appearance = appearance178;
			this.txtProyeccionVentaTotalDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalDiciembre.Enabled = false;
			this.txtProyeccionVentaTotalDiciembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalDiciembre.Location = new System.Drawing.Point(1026, 192);
			this.txtProyeccionVentaTotalDiciembre.Name = "txtProyeccionVentaTotalDiciembre";
			this.txtProyeccionVentaTotalDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalDiciembre.PromptChar = ' ';
			this.txtProyeccionVentaTotalDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalDiciembre.TabIndex = 1282;
			// 
			// txtProyeccionVentaTotalSeptiembre
			// 
			appearance179.BackColorDisabled = System.Drawing.Color.White;
			appearance179.BackColorDisabled2 = System.Drawing.Color.White;
			appearance179.BorderColor = System.Drawing.Color.Black;
			appearance179.BorderColor3DBase = System.Drawing.Color.Black;
			appearance179.ForeColor = System.Drawing.Color.Firebrick;
			appearance179.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalSeptiembre.Appearance = appearance179;
			this.txtProyeccionVentaTotalSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalSeptiembre.Enabled = false;
			this.txtProyeccionVentaTotalSeptiembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalSeptiembre.Location = new System.Drawing.Point(762, 192);
			this.txtProyeccionVentaTotalSeptiembre.Name = "txtProyeccionVentaTotalSeptiembre";
			this.txtProyeccionVentaTotalSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalSeptiembre.PromptChar = ' ';
			this.txtProyeccionVentaTotalSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalSeptiembre.TabIndex = 1281;
			// 
			// txtProyeccionVentaTotalOctubre
			// 
			appearance180.BackColorDisabled = System.Drawing.Color.White;
			appearance180.BackColorDisabled2 = System.Drawing.Color.White;
			appearance180.BorderColor = System.Drawing.Color.Black;
			appearance180.BorderColor3DBase = System.Drawing.Color.Black;
			appearance180.ForeColor = System.Drawing.Color.Firebrick;
			appearance180.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalOctubre.Appearance = appearance180;
			this.txtProyeccionVentaTotalOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalOctubre.Enabled = false;
			this.txtProyeccionVentaTotalOctubre.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalOctubre.Location = new System.Drawing.Point(850, 192);
			this.txtProyeccionVentaTotalOctubre.Name = "txtProyeccionVentaTotalOctubre";
			this.txtProyeccionVentaTotalOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalOctubre.PromptChar = ' ';
			this.txtProyeccionVentaTotalOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalOctubre.TabIndex = 1280;
			// 
			// txtProyeccionVentaTotalJulio
			// 
			appearance181.BackColorDisabled = System.Drawing.Color.White;
			appearance181.BackColorDisabled2 = System.Drawing.Color.White;
			appearance181.BorderColor = System.Drawing.Color.Black;
			appearance181.BorderColor3DBase = System.Drawing.Color.Black;
			appearance181.ForeColor = System.Drawing.Color.Firebrick;
			appearance181.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalJulio.Appearance = appearance181;
			this.txtProyeccionVentaTotalJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalJulio.Enabled = false;
			this.txtProyeccionVentaTotalJulio.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalJulio.Location = new System.Drawing.Point(586, 192);
			this.txtProyeccionVentaTotalJulio.Name = "txtProyeccionVentaTotalJulio";
			this.txtProyeccionVentaTotalJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalJulio.PromptChar = ' ';
			this.txtProyeccionVentaTotalJulio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalJulio.TabIndex = 1279;
			// 
			// txtProyeccionVentaTotalAgosto
			// 
			appearance182.BackColorDisabled = System.Drawing.Color.White;
			appearance182.BackColorDisabled2 = System.Drawing.Color.White;
			appearance182.BorderColor = System.Drawing.Color.Black;
			appearance182.BorderColor3DBase = System.Drawing.Color.Black;
			appearance182.ForeColor = System.Drawing.Color.Firebrick;
			appearance182.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalAgosto.Appearance = appearance182;
			this.txtProyeccionVentaTotalAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalAgosto.Enabled = false;
			this.txtProyeccionVentaTotalAgosto.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalAgosto.Location = new System.Drawing.Point(674, 192);
			this.txtProyeccionVentaTotalAgosto.Name = "txtProyeccionVentaTotalAgosto";
			this.txtProyeccionVentaTotalAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalAgosto.PromptChar = ' ';
			this.txtProyeccionVentaTotalAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalAgosto.TabIndex = 1278;
			// 
			// txtProyeccionVentaTotalMayo
			// 
			appearance183.BackColorDisabled = System.Drawing.Color.White;
			appearance183.BackColorDisabled2 = System.Drawing.Color.White;
			appearance183.BorderColor = System.Drawing.Color.Black;
			appearance183.BorderColor3DBase = System.Drawing.Color.Black;
			appearance183.ForeColor = System.Drawing.Color.Firebrick;
			appearance183.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalMayo.Appearance = appearance183;
			this.txtProyeccionVentaTotalMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalMayo.Enabled = false;
			this.txtProyeccionVentaTotalMayo.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalMayo.Location = new System.Drawing.Point(410, 192);
			this.txtProyeccionVentaTotalMayo.Name = "txtProyeccionVentaTotalMayo";
			this.txtProyeccionVentaTotalMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalMayo.PromptChar = ' ';
			this.txtProyeccionVentaTotalMayo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalMayo.TabIndex = 1277;
			// 
			// txtProyeccionVentaTotalJunio
			// 
			appearance184.BackColorDisabled = System.Drawing.Color.White;
			appearance184.BackColorDisabled2 = System.Drawing.Color.White;
			appearance184.BorderColor = System.Drawing.Color.Black;
			appearance184.BorderColor3DBase = System.Drawing.Color.Black;
			appearance184.ForeColor = System.Drawing.Color.Firebrick;
			appearance184.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalJunio.Appearance = appearance184;
			this.txtProyeccionVentaTotalJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalJunio.Enabled = false;
			this.txtProyeccionVentaTotalJunio.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalJunio.Location = new System.Drawing.Point(498, 192);
			this.txtProyeccionVentaTotalJunio.Name = "txtProyeccionVentaTotalJunio";
			this.txtProyeccionVentaTotalJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalJunio.PromptChar = ' ';
			this.txtProyeccionVentaTotalJunio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalJunio.TabIndex = 1276;
			// 
			// txtProyeccionVentaTotalMarzo
			// 
			appearance185.BackColorDisabled = System.Drawing.Color.White;
			appearance185.BackColorDisabled2 = System.Drawing.Color.White;
			appearance185.BorderColor = System.Drawing.Color.Black;
			appearance185.BorderColor3DBase = System.Drawing.Color.Black;
			appearance185.ForeColor = System.Drawing.Color.Firebrick;
			appearance185.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalMarzo.Appearance = appearance185;
			this.txtProyeccionVentaTotalMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalMarzo.Enabled = false;
			this.txtProyeccionVentaTotalMarzo.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalMarzo.Location = new System.Drawing.Point(234, 192);
			this.txtProyeccionVentaTotalMarzo.Name = "txtProyeccionVentaTotalMarzo";
			this.txtProyeccionVentaTotalMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalMarzo.PromptChar = ' ';
			this.txtProyeccionVentaTotalMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalMarzo.TabIndex = 1275;
			// 
			// txtProyeccionVentaTotalAbril
			// 
			appearance186.BackColorDisabled = System.Drawing.Color.White;
			appearance186.BackColorDisabled2 = System.Drawing.Color.White;
			appearance186.BorderColor = System.Drawing.Color.Black;
			appearance186.BorderColor3DBase = System.Drawing.Color.Black;
			appearance186.ForeColor = System.Drawing.Color.Firebrick;
			appearance186.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalAbril.Appearance = appearance186;
			this.txtProyeccionVentaTotalAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalAbril.Enabled = false;
			this.txtProyeccionVentaTotalAbril.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalAbril.Location = new System.Drawing.Point(322, 192);
			this.txtProyeccionVentaTotalAbril.Name = "txtProyeccionVentaTotalAbril";
			this.txtProyeccionVentaTotalAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalAbril.PromptChar = ' ';
			this.txtProyeccionVentaTotalAbril.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalAbril.TabIndex = 1274;
			// 
			// txtProyeccionVentaTotalFebrero
			// 
			appearance187.BackColorDisabled = System.Drawing.Color.White;
			appearance187.BackColorDisabled2 = System.Drawing.Color.White;
			appearance187.BorderColor = System.Drawing.Color.Black;
			appearance187.BorderColor3DBase = System.Drawing.Color.Black;
			appearance187.ForeColor = System.Drawing.Color.Firebrick;
			appearance187.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalFebrero.Appearance = appearance187;
			this.txtProyeccionVentaTotalFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalFebrero.Enabled = false;
			this.txtProyeccionVentaTotalFebrero.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalFebrero.Location = new System.Drawing.Point(146, 192);
			this.txtProyeccionVentaTotalFebrero.Name = "txtProyeccionVentaTotalFebrero";
			this.txtProyeccionVentaTotalFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalFebrero.PromptChar = ' ';
			this.txtProyeccionVentaTotalFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalFebrero.TabIndex = 1273;
			// 
			// txtProyeccionVentaTotalEnero
			// 
			appearance188.BackColorDisabled = System.Drawing.Color.White;
			appearance188.BackColorDisabled2 = System.Drawing.Color.White;
			appearance188.BorderColor = System.Drawing.Color.Black;
			appearance188.BorderColor3DBase = System.Drawing.Color.Black;
			appearance188.ForeColor = System.Drawing.Color.Firebrick;
			appearance188.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTotalEnero.Appearance = appearance188;
			this.txtProyeccionVentaTotalEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTotalEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTotalEnero.Enabled = false;
			this.txtProyeccionVentaTotalEnero.FormatString = "#,##0.00";
			this.txtProyeccionVentaTotalEnero.Location = new System.Drawing.Point(58, 192);
			this.txtProyeccionVentaTotalEnero.Name = "txtProyeccionVentaTotalEnero";
			this.txtProyeccionVentaTotalEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTotalEnero.PromptChar = ' ';
			this.txtProyeccionVentaTotalEnero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTotalEnero.TabIndex = 1272;
			// 
			// label24
			// 
			this.label24.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label24.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label24.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label24.ForeColor = System.Drawing.Color.Black;
			this.label24.Location = new System.Drawing.Point(0, 192);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(1218, 1);
			this.label24.TabIndex = 1271;
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label2
			// 
			this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Firebrick;
			this.label2.Location = new System.Drawing.Point(2, 216);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(1218, 16);
			this.label2.TabIndex = 1261;
			this.label2.Text = "MARGEN";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label1
			// 
			this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(2, 80);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(1218, 16);
			this.label1.TabIndex = 1260;
			this.label1.Text = "VENTAS";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtProyeccionMargenCredito
			// 
			appearance189.BackColorDisabled = System.Drawing.Color.White;
			appearance189.BackColorDisabled2 = System.Drawing.Color.White;
			appearance189.BorderColor = System.Drawing.Color.Black;
			appearance189.BorderColor3DBase = System.Drawing.Color.Black;
			appearance189.FontData.BoldAsString = "True";
			appearance189.ForeColor = System.Drawing.Color.Firebrick;
			appearance189.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenCredito.Appearance = appearance189;
			this.txtProyeccionMargenCredito.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCredito.Enabled = false;
			this.txtProyeccionMargenCredito.FormatString = "#,##0.00";
			this.txtProyeccionMargenCredito.Location = new System.Drawing.Point(1114, 304);
			this.txtProyeccionMargenCredito.Name = "txtProyeccionMargenCredito";
			this.txtProyeccionMargenCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCredito.PromptChar = ' ';
			this.txtProyeccionMargenCredito.Size = new System.Drawing.Size(100, 19);
			this.txtProyeccionMargenCredito.TabIndex = 1259;
			// 
			// txtProyeccionMargenCreditoNoviembre
			// 
			appearance190.BackColorDisabled = System.Drawing.Color.White;
			appearance190.BackColorDisabled2 = System.Drawing.Color.White;
			appearance190.BorderColor = System.Drawing.Color.Black;
			appearance190.BorderColor3DBase = System.Drawing.Color.Black;
			appearance190.ForeColor = System.Drawing.Color.Black;
			appearance190.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoNoviembre.Appearance = appearance190;
			this.txtProyeccionMargenCreditoNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoNoviembre.Enabled = false;
			this.txtProyeccionMargenCreditoNoviembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoNoviembre.Location = new System.Drawing.Point(938, 304);
			this.txtProyeccionMargenCreditoNoviembre.Name = "txtProyeccionMargenCreditoNoviembre";
			this.txtProyeccionMargenCreditoNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoNoviembre.PromptChar = ' ';
			this.txtProyeccionMargenCreditoNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoNoviembre.TabIndex = 1258;
			// 
			// txtProyeccionMargenCreditoDiciembre
			// 
			appearance191.BackColorDisabled = System.Drawing.Color.White;
			appearance191.BackColorDisabled2 = System.Drawing.Color.White;
			appearance191.BorderColor = System.Drawing.Color.Black;
			appearance191.BorderColor3DBase = System.Drawing.Color.Black;
			appearance191.ForeColor = System.Drawing.Color.Black;
			appearance191.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoDiciembre.Appearance = appearance191;
			this.txtProyeccionMargenCreditoDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoDiciembre.Enabled = false;
			this.txtProyeccionMargenCreditoDiciembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoDiciembre.Location = new System.Drawing.Point(1026, 304);
			this.txtProyeccionMargenCreditoDiciembre.Name = "txtProyeccionMargenCreditoDiciembre";
			this.txtProyeccionMargenCreditoDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoDiciembre.PromptChar = ' ';
			this.txtProyeccionMargenCreditoDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoDiciembre.TabIndex = 1257;
			// 
			// txtProyeccionMargenCreditoSeptiembre
			// 
			appearance192.BackColorDisabled = System.Drawing.Color.White;
			appearance192.BackColorDisabled2 = System.Drawing.Color.White;
			appearance192.BorderColor = System.Drawing.Color.Black;
			appearance192.BorderColor3DBase = System.Drawing.Color.Black;
			appearance192.ForeColor = System.Drawing.Color.Black;
			appearance192.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoSeptiembre.Appearance = appearance192;
			this.txtProyeccionMargenCreditoSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoSeptiembre.Enabled = false;
			this.txtProyeccionMargenCreditoSeptiembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoSeptiembre.Location = new System.Drawing.Point(762, 304);
			this.txtProyeccionMargenCreditoSeptiembre.Name = "txtProyeccionMargenCreditoSeptiembre";
			this.txtProyeccionMargenCreditoSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoSeptiembre.PromptChar = ' ';
			this.txtProyeccionMargenCreditoSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoSeptiembre.TabIndex = 1256;
			// 
			// txtProyeccionMargenCreditoOctubre
			// 
			appearance193.BackColorDisabled = System.Drawing.Color.White;
			appearance193.BackColorDisabled2 = System.Drawing.Color.White;
			appearance193.BorderColor = System.Drawing.Color.Black;
			appearance193.BorderColor3DBase = System.Drawing.Color.Black;
			appearance193.ForeColor = System.Drawing.Color.Black;
			appearance193.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoOctubre.Appearance = appearance193;
			this.txtProyeccionMargenCreditoOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoOctubre.Enabled = false;
			this.txtProyeccionMargenCreditoOctubre.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoOctubre.Location = new System.Drawing.Point(850, 304);
			this.txtProyeccionMargenCreditoOctubre.Name = "txtProyeccionMargenCreditoOctubre";
			this.txtProyeccionMargenCreditoOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoOctubre.PromptChar = ' ';
			this.txtProyeccionMargenCreditoOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoOctubre.TabIndex = 1255;
			// 
			// txtProyeccionMargenCreditoJulio
			// 
			appearance194.BackColorDisabled = System.Drawing.Color.White;
			appearance194.BackColorDisabled2 = System.Drawing.Color.White;
			appearance194.BorderColor = System.Drawing.Color.Black;
			appearance194.BorderColor3DBase = System.Drawing.Color.Black;
			appearance194.ForeColor = System.Drawing.Color.Black;
			appearance194.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoJulio.Appearance = appearance194;
			this.txtProyeccionMargenCreditoJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoJulio.Enabled = false;
			this.txtProyeccionMargenCreditoJulio.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoJulio.Location = new System.Drawing.Point(586, 304);
			this.txtProyeccionMargenCreditoJulio.Name = "txtProyeccionMargenCreditoJulio";
			this.txtProyeccionMargenCreditoJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoJulio.PromptChar = ' ';
			this.txtProyeccionMargenCreditoJulio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoJulio.TabIndex = 1254;
			// 
			// txtProyeccionMargenCreditoAgosto
			// 
			appearance195.BackColorDisabled = System.Drawing.Color.White;
			appearance195.BackColorDisabled2 = System.Drawing.Color.White;
			appearance195.BorderColor = System.Drawing.Color.Black;
			appearance195.BorderColor3DBase = System.Drawing.Color.Black;
			appearance195.ForeColor = System.Drawing.Color.Black;
			appearance195.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoAgosto.Appearance = appearance195;
			this.txtProyeccionMargenCreditoAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoAgosto.Enabled = false;
			this.txtProyeccionMargenCreditoAgosto.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoAgosto.Location = new System.Drawing.Point(674, 304);
			this.txtProyeccionMargenCreditoAgosto.Name = "txtProyeccionMargenCreditoAgosto";
			this.txtProyeccionMargenCreditoAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoAgosto.PromptChar = ' ';
			this.txtProyeccionMargenCreditoAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoAgosto.TabIndex = 1253;
			// 
			// txtProyeccionMargenCreditoMayo
			// 
			appearance196.BackColorDisabled = System.Drawing.Color.White;
			appearance196.BackColorDisabled2 = System.Drawing.Color.White;
			appearance196.BorderColor = System.Drawing.Color.Black;
			appearance196.BorderColor3DBase = System.Drawing.Color.Black;
			appearance196.ForeColor = System.Drawing.Color.Black;
			appearance196.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoMayo.Appearance = appearance196;
			this.txtProyeccionMargenCreditoMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoMayo.Enabled = false;
			this.txtProyeccionMargenCreditoMayo.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoMayo.Location = new System.Drawing.Point(410, 304);
			this.txtProyeccionMargenCreditoMayo.Name = "txtProyeccionMargenCreditoMayo";
			this.txtProyeccionMargenCreditoMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoMayo.PromptChar = ' ';
			this.txtProyeccionMargenCreditoMayo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoMayo.TabIndex = 1252;
			// 
			// txtProyeccionMargenCreditoJunio
			// 
			appearance197.BackColorDisabled = System.Drawing.Color.White;
			appearance197.BackColorDisabled2 = System.Drawing.Color.White;
			appearance197.BorderColor = System.Drawing.Color.Black;
			appearance197.BorderColor3DBase = System.Drawing.Color.Black;
			appearance197.ForeColor = System.Drawing.Color.Black;
			appearance197.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoJunio.Appearance = appearance197;
			this.txtProyeccionMargenCreditoJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoJunio.Enabled = false;
			this.txtProyeccionMargenCreditoJunio.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoJunio.Location = new System.Drawing.Point(498, 304);
			this.txtProyeccionMargenCreditoJunio.Name = "txtProyeccionMargenCreditoJunio";
			this.txtProyeccionMargenCreditoJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoJunio.PromptChar = ' ';
			this.txtProyeccionMargenCreditoJunio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoJunio.TabIndex = 1251;
			// 
			// txtProyeccionMargenCreditoMarzo
			// 
			appearance198.BackColorDisabled = System.Drawing.Color.White;
			appearance198.BackColorDisabled2 = System.Drawing.Color.White;
			appearance198.BorderColor = System.Drawing.Color.Black;
			appearance198.BorderColor3DBase = System.Drawing.Color.Black;
			appearance198.ForeColor = System.Drawing.Color.Black;
			appearance198.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoMarzo.Appearance = appearance198;
			this.txtProyeccionMargenCreditoMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoMarzo.Enabled = false;
			this.txtProyeccionMargenCreditoMarzo.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoMarzo.Location = new System.Drawing.Point(234, 304);
			this.txtProyeccionMargenCreditoMarzo.Name = "txtProyeccionMargenCreditoMarzo";
			this.txtProyeccionMargenCreditoMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoMarzo.PromptChar = ' ';
			this.txtProyeccionMargenCreditoMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoMarzo.TabIndex = 1250;
			// 
			// txtProyeccionMargenCreditoAbril
			// 
			appearance199.BackColorDisabled = System.Drawing.Color.White;
			appearance199.BackColorDisabled2 = System.Drawing.Color.White;
			appearance199.BorderColor = System.Drawing.Color.Black;
			appearance199.BorderColor3DBase = System.Drawing.Color.Black;
			appearance199.ForeColor = System.Drawing.Color.Black;
			appearance199.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoAbril.Appearance = appearance199;
			this.txtProyeccionMargenCreditoAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoAbril.Enabled = false;
			this.txtProyeccionMargenCreditoAbril.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoAbril.Location = new System.Drawing.Point(322, 304);
			this.txtProyeccionMargenCreditoAbril.Name = "txtProyeccionMargenCreditoAbril";
			this.txtProyeccionMargenCreditoAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoAbril.PromptChar = ' ';
			this.txtProyeccionMargenCreditoAbril.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoAbril.TabIndex = 1249;
			// 
			// txtProyeccionMargenCreditoEnero
			// 
			appearance200.BackColorDisabled = System.Drawing.Color.White;
			appearance200.BackColorDisabled2 = System.Drawing.Color.White;
			appearance200.BorderColor = System.Drawing.Color.Black;
			appearance200.BorderColor3DBase = System.Drawing.Color.Black;
			appearance200.ForeColor = System.Drawing.Color.Black;
			appearance200.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoEnero.Appearance = appearance200;
			this.txtProyeccionMargenCreditoEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoEnero.Enabled = false;
			this.txtProyeccionMargenCreditoEnero.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoEnero.Location = new System.Drawing.Point(58, 304);
			this.txtProyeccionMargenCreditoEnero.Name = "txtProyeccionMargenCreditoEnero";
			this.txtProyeccionMargenCreditoEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoEnero.PromptChar = ' ';
			this.txtProyeccionMargenCreditoEnero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoEnero.TabIndex = 1248;
			// 
			// txtProyeccionMargenCreditoFebrero
			// 
			appearance201.BackColorDisabled = System.Drawing.Color.White;
			appearance201.BackColorDisabled2 = System.Drawing.Color.White;
			appearance201.BorderColor = System.Drawing.Color.Black;
			appearance201.BorderColor3DBase = System.Drawing.Color.Black;
			appearance201.ForeColor = System.Drawing.Color.Black;
			appearance201.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenCreditoFebrero.Appearance = appearance201;
			this.txtProyeccionMargenCreditoFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenCreditoFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenCreditoFebrero.Enabled = false;
			this.txtProyeccionMargenCreditoFebrero.FormatString = "#,##0.00";
			this.txtProyeccionMargenCreditoFebrero.Location = new System.Drawing.Point(146, 304);
			this.txtProyeccionMargenCreditoFebrero.Name = "txtProyeccionMargenCreditoFebrero";
			this.txtProyeccionMargenCreditoFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenCreditoFebrero.PromptChar = ' ';
			this.txtProyeccionMargenCreditoFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenCreditoFebrero.TabIndex = 1247;
			// 
			// txtProyeccionMargenTarjeta
			// 
			appearance202.BackColorDisabled = System.Drawing.Color.White;
			appearance202.BackColorDisabled2 = System.Drawing.Color.White;
			appearance202.BorderColor = System.Drawing.Color.Black;
			appearance202.BorderColor3DBase = System.Drawing.Color.Black;
			appearance202.FontData.BoldAsString = "True";
			appearance202.ForeColor = System.Drawing.Color.Firebrick;
			appearance202.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenTarjeta.Appearance = appearance202;
			this.txtProyeccionMargenTarjeta.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjeta.Enabled = false;
			this.txtProyeccionMargenTarjeta.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjeta.Location = new System.Drawing.Point(1114, 280);
			this.txtProyeccionMargenTarjeta.Name = "txtProyeccionMargenTarjeta";
			this.txtProyeccionMargenTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjeta.PromptChar = ' ';
			this.txtProyeccionMargenTarjeta.Size = new System.Drawing.Size(100, 19);
			this.txtProyeccionMargenTarjeta.TabIndex = 1246;
			// 
			// txtProyeccionMargenTarjetaNoviembre
			// 
			appearance203.BackColorDisabled = System.Drawing.Color.White;
			appearance203.BackColorDisabled2 = System.Drawing.Color.White;
			appearance203.BorderColor = System.Drawing.Color.Black;
			appearance203.BorderColor3DBase = System.Drawing.Color.Black;
			appearance203.ForeColor = System.Drawing.Color.Black;
			appearance203.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaNoviembre.Appearance = appearance203;
			this.txtProyeccionMargenTarjetaNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaNoviembre.Enabled = false;
			this.txtProyeccionMargenTarjetaNoviembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaNoviembre.Location = new System.Drawing.Point(938, 280);
			this.txtProyeccionMargenTarjetaNoviembre.Name = "txtProyeccionMargenTarjetaNoviembre";
			this.txtProyeccionMargenTarjetaNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaNoviembre.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaNoviembre.TabIndex = 1245;
			// 
			// txtProyeccionMargenTarjetaDiciembre
			// 
			appearance204.BackColorDisabled = System.Drawing.Color.White;
			appearance204.BackColorDisabled2 = System.Drawing.Color.White;
			appearance204.BorderColor = System.Drawing.Color.Black;
			appearance204.BorderColor3DBase = System.Drawing.Color.Black;
			appearance204.ForeColor = System.Drawing.Color.Black;
			appearance204.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaDiciembre.Appearance = appearance204;
			this.txtProyeccionMargenTarjetaDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaDiciembre.Enabled = false;
			this.txtProyeccionMargenTarjetaDiciembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaDiciembre.Location = new System.Drawing.Point(1026, 280);
			this.txtProyeccionMargenTarjetaDiciembre.Name = "txtProyeccionMargenTarjetaDiciembre";
			this.txtProyeccionMargenTarjetaDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaDiciembre.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaDiciembre.TabIndex = 1244;
			// 
			// txtProyeccionMargenTarjetaSeptiembre
			// 
			appearance205.BackColorDisabled = System.Drawing.Color.White;
			appearance205.BackColorDisabled2 = System.Drawing.Color.White;
			appearance205.BorderColor = System.Drawing.Color.Black;
			appearance205.BorderColor3DBase = System.Drawing.Color.Black;
			appearance205.ForeColor = System.Drawing.Color.Black;
			appearance205.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaSeptiembre.Appearance = appearance205;
			this.txtProyeccionMargenTarjetaSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaSeptiembre.Enabled = false;
			this.txtProyeccionMargenTarjetaSeptiembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaSeptiembre.Location = new System.Drawing.Point(762, 280);
			this.txtProyeccionMargenTarjetaSeptiembre.Name = "txtProyeccionMargenTarjetaSeptiembre";
			this.txtProyeccionMargenTarjetaSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaSeptiembre.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaSeptiembre.TabIndex = 1243;
			// 
			// txtProyeccionMargenTarjetaOctubre
			// 
			appearance206.BackColorDisabled = System.Drawing.Color.White;
			appearance206.BackColorDisabled2 = System.Drawing.Color.White;
			appearance206.BorderColor = System.Drawing.Color.Black;
			appearance206.BorderColor3DBase = System.Drawing.Color.Black;
			appearance206.ForeColor = System.Drawing.Color.Black;
			appearance206.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaOctubre.Appearance = appearance206;
			this.txtProyeccionMargenTarjetaOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaOctubre.Enabled = false;
			this.txtProyeccionMargenTarjetaOctubre.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaOctubre.Location = new System.Drawing.Point(850, 280);
			this.txtProyeccionMargenTarjetaOctubre.Name = "txtProyeccionMargenTarjetaOctubre";
			this.txtProyeccionMargenTarjetaOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaOctubre.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaOctubre.TabIndex = 1242;
			// 
			// txtProyeccionMargenTarjetaJulio
			// 
			appearance207.BackColorDisabled = System.Drawing.Color.White;
			appearance207.BackColorDisabled2 = System.Drawing.Color.White;
			appearance207.BorderColor = System.Drawing.Color.Black;
			appearance207.BorderColor3DBase = System.Drawing.Color.Black;
			appearance207.ForeColor = System.Drawing.Color.Black;
			appearance207.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaJulio.Appearance = appearance207;
			this.txtProyeccionMargenTarjetaJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaJulio.Enabled = false;
			this.txtProyeccionMargenTarjetaJulio.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaJulio.Location = new System.Drawing.Point(586, 280);
			this.txtProyeccionMargenTarjetaJulio.Name = "txtProyeccionMargenTarjetaJulio";
			this.txtProyeccionMargenTarjetaJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaJulio.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaJulio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaJulio.TabIndex = 1241;
			// 
			// txtProyeccionMargenTarjetaAgosto
			// 
			appearance208.BackColorDisabled = System.Drawing.Color.White;
			appearance208.BackColorDisabled2 = System.Drawing.Color.White;
			appearance208.BorderColor = System.Drawing.Color.Black;
			appearance208.BorderColor3DBase = System.Drawing.Color.Black;
			appearance208.ForeColor = System.Drawing.Color.Black;
			appearance208.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaAgosto.Appearance = appearance208;
			this.txtProyeccionMargenTarjetaAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaAgosto.Enabled = false;
			this.txtProyeccionMargenTarjetaAgosto.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaAgosto.Location = new System.Drawing.Point(674, 280);
			this.txtProyeccionMargenTarjetaAgosto.Name = "txtProyeccionMargenTarjetaAgosto";
			this.txtProyeccionMargenTarjetaAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaAgosto.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaAgosto.TabIndex = 1240;
			// 
			// txtProyeccionMargenTarjetaMayo
			// 
			appearance209.BackColorDisabled = System.Drawing.Color.White;
			appearance209.BackColorDisabled2 = System.Drawing.Color.White;
			appearance209.BorderColor = System.Drawing.Color.Black;
			appearance209.BorderColor3DBase = System.Drawing.Color.Black;
			appearance209.ForeColor = System.Drawing.Color.Black;
			appearance209.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaMayo.Appearance = appearance209;
			this.txtProyeccionMargenTarjetaMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaMayo.Enabled = false;
			this.txtProyeccionMargenTarjetaMayo.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaMayo.Location = new System.Drawing.Point(410, 280);
			this.txtProyeccionMargenTarjetaMayo.Name = "txtProyeccionMargenTarjetaMayo";
			this.txtProyeccionMargenTarjetaMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaMayo.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaMayo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaMayo.TabIndex = 1239;
			// 
			// txtProyeccionMargenTarjetaJunio
			// 
			appearance210.BackColorDisabled = System.Drawing.Color.White;
			appearance210.BackColorDisabled2 = System.Drawing.Color.White;
			appearance210.BorderColor = System.Drawing.Color.Black;
			appearance210.BorderColor3DBase = System.Drawing.Color.Black;
			appearance210.ForeColor = System.Drawing.Color.Black;
			appearance210.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaJunio.Appearance = appearance210;
			this.txtProyeccionMargenTarjetaJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaJunio.Enabled = false;
			this.txtProyeccionMargenTarjetaJunio.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaJunio.Location = new System.Drawing.Point(498, 280);
			this.txtProyeccionMargenTarjetaJunio.Name = "txtProyeccionMargenTarjetaJunio";
			this.txtProyeccionMargenTarjetaJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaJunio.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaJunio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaJunio.TabIndex = 1238;
			// 
			// txtProyeccionMargenTarjetaMarzo
			// 
			appearance211.BackColorDisabled = System.Drawing.Color.White;
			appearance211.BackColorDisabled2 = System.Drawing.Color.White;
			appearance211.BorderColor = System.Drawing.Color.Black;
			appearance211.BorderColor3DBase = System.Drawing.Color.Black;
			appearance211.ForeColor = System.Drawing.Color.Black;
			appearance211.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaMarzo.Appearance = appearance211;
			this.txtProyeccionMargenTarjetaMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaMarzo.Enabled = false;
			this.txtProyeccionMargenTarjetaMarzo.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaMarzo.Location = new System.Drawing.Point(234, 280);
			this.txtProyeccionMargenTarjetaMarzo.Name = "txtProyeccionMargenTarjetaMarzo";
			this.txtProyeccionMargenTarjetaMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaMarzo.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaMarzo.TabIndex = 1237;
			// 
			// txtProyeccionMargenTarjetaAbril
			// 
			appearance212.BackColorDisabled = System.Drawing.Color.White;
			appearance212.BackColorDisabled2 = System.Drawing.Color.White;
			appearance212.BorderColor = System.Drawing.Color.Black;
			appearance212.BorderColor3DBase = System.Drawing.Color.Black;
			appearance212.ForeColor = System.Drawing.Color.Black;
			appearance212.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaAbril.Appearance = appearance212;
			this.txtProyeccionMargenTarjetaAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaAbril.Enabled = false;
			this.txtProyeccionMargenTarjetaAbril.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaAbril.Location = new System.Drawing.Point(322, 280);
			this.txtProyeccionMargenTarjetaAbril.Name = "txtProyeccionMargenTarjetaAbril";
			this.txtProyeccionMargenTarjetaAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaAbril.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaAbril.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaAbril.TabIndex = 1236;
			// 
			// txtProyeccionMargenTarjetaEnero
			// 
			appearance213.BackColorDisabled = System.Drawing.Color.White;
			appearance213.BackColorDisabled2 = System.Drawing.Color.White;
			appearance213.BorderColor = System.Drawing.Color.Black;
			appearance213.BorderColor3DBase = System.Drawing.Color.Black;
			appearance213.ForeColor = System.Drawing.Color.Black;
			appearance213.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaEnero.Appearance = appearance213;
			this.txtProyeccionMargenTarjetaEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaEnero.Enabled = false;
			this.txtProyeccionMargenTarjetaEnero.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaEnero.Location = new System.Drawing.Point(58, 280);
			this.txtProyeccionMargenTarjetaEnero.Name = "txtProyeccionMargenTarjetaEnero";
			this.txtProyeccionMargenTarjetaEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaEnero.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaEnero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaEnero.TabIndex = 1235;
			// 
			// txtProyeccionMargenTarjetaFebrero
			// 
			appearance214.BackColorDisabled = System.Drawing.Color.White;
			appearance214.BackColorDisabled2 = System.Drawing.Color.White;
			appearance214.BorderColor = System.Drawing.Color.Black;
			appearance214.BorderColor3DBase = System.Drawing.Color.Black;
			appearance214.ForeColor = System.Drawing.Color.Black;
			appearance214.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenTarjetaFebrero.Appearance = appearance214;
			this.txtProyeccionMargenTarjetaFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenTarjetaFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenTarjetaFebrero.Enabled = false;
			this.txtProyeccionMargenTarjetaFebrero.FormatString = "#,##0.00";
			this.txtProyeccionMargenTarjetaFebrero.Location = new System.Drawing.Point(146, 280);
			this.txtProyeccionMargenTarjetaFebrero.Name = "txtProyeccionMargenTarjetaFebrero";
			this.txtProyeccionMargenTarjetaFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenTarjetaFebrero.PromptChar = ' ';
			this.txtProyeccionMargenTarjetaFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenTarjetaFebrero.TabIndex = 1234;
			// 
			// txtProyeccionMargenEfectivo
			// 
			appearance215.BackColorDisabled = System.Drawing.Color.White;
			appearance215.BackColorDisabled2 = System.Drawing.Color.White;
			appearance215.BorderColor = System.Drawing.Color.Black;
			appearance215.BorderColor3DBase = System.Drawing.Color.Black;
			appearance215.FontData.BoldAsString = "True";
			appearance215.ForeColor = System.Drawing.Color.Firebrick;
			appearance215.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenEfectivo.Appearance = appearance215;
			this.txtProyeccionMargenEfectivo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivo.Enabled = false;
			this.txtProyeccionMargenEfectivo.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivo.Location = new System.Drawing.Point(1114, 256);
			this.txtProyeccionMargenEfectivo.Name = "txtProyeccionMargenEfectivo";
			this.txtProyeccionMargenEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivo.PromptChar = ' ';
			this.txtProyeccionMargenEfectivo.Size = new System.Drawing.Size(100, 19);
			this.txtProyeccionMargenEfectivo.TabIndex = 1233;
			// 
			// txtProyeccionMargenEfectivoNoviembre
			// 
			appearance216.BackColorDisabled = System.Drawing.Color.White;
			appearance216.BackColorDisabled2 = System.Drawing.Color.White;
			appearance216.BorderColor = System.Drawing.Color.Black;
			appearance216.BorderColor3DBase = System.Drawing.Color.Black;
			appearance216.ForeColor = System.Drawing.Color.Black;
			appearance216.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoNoviembre.Appearance = appearance216;
			this.txtProyeccionMargenEfectivoNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoNoviembre.Enabled = false;
			this.txtProyeccionMargenEfectivoNoviembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoNoviembre.Location = new System.Drawing.Point(938, 256);
			this.txtProyeccionMargenEfectivoNoviembre.Name = "txtProyeccionMargenEfectivoNoviembre";
			this.txtProyeccionMargenEfectivoNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoNoviembre.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoNoviembre.TabIndex = 1232;
			// 
			// txtProyeccionMargenEfectivoDiciembre
			// 
			appearance217.BackColorDisabled = System.Drawing.Color.White;
			appearance217.BackColorDisabled2 = System.Drawing.Color.White;
			appearance217.BorderColor = System.Drawing.Color.Black;
			appearance217.BorderColor3DBase = System.Drawing.Color.Black;
			appearance217.ForeColor = System.Drawing.Color.Black;
			appearance217.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoDiciembre.Appearance = appearance217;
			this.txtProyeccionMargenEfectivoDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoDiciembre.Enabled = false;
			this.txtProyeccionMargenEfectivoDiciembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoDiciembre.Location = new System.Drawing.Point(1026, 256);
			this.txtProyeccionMargenEfectivoDiciembre.Name = "txtProyeccionMargenEfectivoDiciembre";
			this.txtProyeccionMargenEfectivoDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoDiciembre.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoDiciembre.TabIndex = 1231;
			// 
			// txtProyeccionMargenEfectivoSeptiembre
			// 
			appearance218.BackColorDisabled = System.Drawing.Color.White;
			appearance218.BackColorDisabled2 = System.Drawing.Color.White;
			appearance218.BorderColor = System.Drawing.Color.Black;
			appearance218.BorderColor3DBase = System.Drawing.Color.Black;
			appearance218.ForeColor = System.Drawing.Color.Black;
			appearance218.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoSeptiembre.Appearance = appearance218;
			this.txtProyeccionMargenEfectivoSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoSeptiembre.Enabled = false;
			this.txtProyeccionMargenEfectivoSeptiembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoSeptiembre.Location = new System.Drawing.Point(762, 256);
			this.txtProyeccionMargenEfectivoSeptiembre.Name = "txtProyeccionMargenEfectivoSeptiembre";
			this.txtProyeccionMargenEfectivoSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoSeptiembre.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoSeptiembre.TabIndex = 1230;
			// 
			// txtProyeccionMargenEfectivoOctubre
			// 
			appearance219.BackColorDisabled = System.Drawing.Color.White;
			appearance219.BackColorDisabled2 = System.Drawing.Color.White;
			appearance219.BorderColor = System.Drawing.Color.Black;
			appearance219.BorderColor3DBase = System.Drawing.Color.Black;
			appearance219.ForeColor = System.Drawing.Color.Black;
			appearance219.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoOctubre.Appearance = appearance219;
			this.txtProyeccionMargenEfectivoOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoOctubre.Enabled = false;
			this.txtProyeccionMargenEfectivoOctubre.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoOctubre.Location = new System.Drawing.Point(850, 256);
			this.txtProyeccionMargenEfectivoOctubre.Name = "txtProyeccionMargenEfectivoOctubre";
			this.txtProyeccionMargenEfectivoOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoOctubre.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoOctubre.TabIndex = 1229;
			// 
			// txtProyeccionMargenEfectivoJulio
			// 
			appearance220.BackColorDisabled = System.Drawing.Color.White;
			appearance220.BackColorDisabled2 = System.Drawing.Color.White;
			appearance220.BorderColor = System.Drawing.Color.Black;
			appearance220.BorderColor3DBase = System.Drawing.Color.Black;
			appearance220.ForeColor = System.Drawing.Color.Black;
			appearance220.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoJulio.Appearance = appearance220;
			this.txtProyeccionMargenEfectivoJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoJulio.Enabled = false;
			this.txtProyeccionMargenEfectivoJulio.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoJulio.Location = new System.Drawing.Point(586, 256);
			this.txtProyeccionMargenEfectivoJulio.Name = "txtProyeccionMargenEfectivoJulio";
			this.txtProyeccionMargenEfectivoJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoJulio.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoJulio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoJulio.TabIndex = 1228;
			// 
			// txtProyeccionMargenEfectivoAgosto
			// 
			appearance221.BackColorDisabled = System.Drawing.Color.White;
			appearance221.BackColorDisabled2 = System.Drawing.Color.White;
			appearance221.BorderColor = System.Drawing.Color.Black;
			appearance221.BorderColor3DBase = System.Drawing.Color.Black;
			appearance221.ForeColor = System.Drawing.Color.Black;
			appearance221.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoAgosto.Appearance = appearance221;
			this.txtProyeccionMargenEfectivoAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoAgosto.Enabled = false;
			this.txtProyeccionMargenEfectivoAgosto.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoAgosto.Location = new System.Drawing.Point(674, 256);
			this.txtProyeccionMargenEfectivoAgosto.Name = "txtProyeccionMargenEfectivoAgosto";
			this.txtProyeccionMargenEfectivoAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoAgosto.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoAgosto.TabIndex = 1227;
			// 
			// txtProyeccionMargenEfectivoMayo
			// 
			appearance222.BackColorDisabled = System.Drawing.Color.White;
			appearance222.BackColorDisabled2 = System.Drawing.Color.White;
			appearance222.BorderColor = System.Drawing.Color.Black;
			appearance222.BorderColor3DBase = System.Drawing.Color.Black;
			appearance222.ForeColor = System.Drawing.Color.Black;
			appearance222.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoMayo.Appearance = appearance222;
			this.txtProyeccionMargenEfectivoMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoMayo.Enabled = false;
			this.txtProyeccionMargenEfectivoMayo.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoMayo.Location = new System.Drawing.Point(410, 256);
			this.txtProyeccionMargenEfectivoMayo.Name = "txtProyeccionMargenEfectivoMayo";
			this.txtProyeccionMargenEfectivoMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoMayo.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoMayo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoMayo.TabIndex = 1226;
			// 
			// txtProyeccionMargenEfectivoJunio
			// 
			appearance223.BackColorDisabled = System.Drawing.Color.White;
			appearance223.BackColorDisabled2 = System.Drawing.Color.White;
			appearance223.BorderColor = System.Drawing.Color.Black;
			appearance223.BorderColor3DBase = System.Drawing.Color.Black;
			appearance223.ForeColor = System.Drawing.Color.Black;
			appearance223.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoJunio.Appearance = appearance223;
			this.txtProyeccionMargenEfectivoJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoJunio.Enabled = false;
			this.txtProyeccionMargenEfectivoJunio.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoJunio.Location = new System.Drawing.Point(498, 256);
			this.txtProyeccionMargenEfectivoJunio.Name = "txtProyeccionMargenEfectivoJunio";
			this.txtProyeccionMargenEfectivoJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoJunio.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoJunio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoJunio.TabIndex = 1225;
			// 
			// txtProyeccionMargenEfectivoMarzo
			// 
			appearance224.BackColorDisabled = System.Drawing.Color.White;
			appearance224.BackColorDisabled2 = System.Drawing.Color.White;
			appearance224.BorderColor = System.Drawing.Color.Black;
			appearance224.BorderColor3DBase = System.Drawing.Color.Black;
			appearance224.ForeColor = System.Drawing.Color.Black;
			appearance224.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoMarzo.Appearance = appearance224;
			this.txtProyeccionMargenEfectivoMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoMarzo.Enabled = false;
			this.txtProyeccionMargenEfectivoMarzo.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoMarzo.Location = new System.Drawing.Point(234, 256);
			this.txtProyeccionMargenEfectivoMarzo.Name = "txtProyeccionMargenEfectivoMarzo";
			this.txtProyeccionMargenEfectivoMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoMarzo.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoMarzo.TabIndex = 1224;
			// 
			// txtProyeccionMargenEfectivoAbril
			// 
			appearance225.BackColorDisabled = System.Drawing.Color.White;
			appearance225.BackColorDisabled2 = System.Drawing.Color.White;
			appearance225.BorderColor = System.Drawing.Color.Black;
			appearance225.BorderColor3DBase = System.Drawing.Color.Black;
			appearance225.ForeColor = System.Drawing.Color.Black;
			appearance225.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoAbril.Appearance = appearance225;
			this.txtProyeccionMargenEfectivoAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoAbril.Enabled = false;
			this.txtProyeccionMargenEfectivoAbril.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoAbril.Location = new System.Drawing.Point(322, 256);
			this.txtProyeccionMargenEfectivoAbril.Name = "txtProyeccionMargenEfectivoAbril";
			this.txtProyeccionMargenEfectivoAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoAbril.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoAbril.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoAbril.TabIndex = 1223;
			// 
			// txtProyeccionMargenEfectivoEnero
			// 
			appearance226.BackColorDisabled = System.Drawing.Color.White;
			appearance226.BackColorDisabled2 = System.Drawing.Color.White;
			appearance226.BorderColor = System.Drawing.Color.Black;
			appearance226.BorderColor3DBase = System.Drawing.Color.Black;
			appearance226.ForeColor = System.Drawing.Color.Black;
			appearance226.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoEnero.Appearance = appearance226;
			this.txtProyeccionMargenEfectivoEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoEnero.Enabled = false;
			this.txtProyeccionMargenEfectivoEnero.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoEnero.Location = new System.Drawing.Point(58, 256);
			this.txtProyeccionMargenEfectivoEnero.Name = "txtProyeccionMargenEfectivoEnero";
			this.txtProyeccionMargenEfectivoEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoEnero.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoEnero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoEnero.TabIndex = 1222;
			// 
			// txtProyeccionMargenEfectivoFebrero
			// 
			appearance227.BackColorDisabled = System.Drawing.Color.White;
			appearance227.BackColorDisabled2 = System.Drawing.Color.White;
			appearance227.BorderColor = System.Drawing.Color.Black;
			appearance227.BorderColor3DBase = System.Drawing.Color.Black;
			appearance227.ForeColor = System.Drawing.Color.Black;
			appearance227.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenEfectivoFebrero.Appearance = appearance227;
			this.txtProyeccionMargenEfectivoFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenEfectivoFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenEfectivoFebrero.Enabled = false;
			this.txtProyeccionMargenEfectivoFebrero.FormatString = "#,##0.00";
			this.txtProyeccionMargenEfectivoFebrero.Location = new System.Drawing.Point(146, 256);
			this.txtProyeccionMargenEfectivoFebrero.Name = "txtProyeccionMargenEfectivoFebrero";
			this.txtProyeccionMargenEfectivoFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenEfectivoFebrero.PromptChar = ' ';
			this.txtProyeccionMargenEfectivoFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenEfectivoFebrero.TabIndex = 1221;
			// 
			// txtProyeccionMargenDistribuidor
			// 
			appearance228.BackColorDisabled = System.Drawing.Color.White;
			appearance228.BackColorDisabled2 = System.Drawing.Color.White;
			appearance228.BorderColor = System.Drawing.Color.Black;
			appearance228.BorderColor3DBase = System.Drawing.Color.Black;
			appearance228.FontData.BoldAsString = "True";
			appearance228.ForeColor = System.Drawing.Color.Firebrick;
			appearance228.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMargenDistribuidor.Appearance = appearance228;
			this.txtProyeccionMargenDistribuidor.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidor.Enabled = false;
			this.txtProyeccionMargenDistribuidor.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidor.Location = new System.Drawing.Point(1114, 232);
			this.txtProyeccionMargenDistribuidor.Name = "txtProyeccionMargenDistribuidor";
			this.txtProyeccionMargenDistribuidor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidor.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidor.Size = new System.Drawing.Size(100, 19);
			this.txtProyeccionMargenDistribuidor.TabIndex = 1220;
			// 
			// txtProyeccionMargenDistribuidorNoviembre
			// 
			appearance229.BackColorDisabled = System.Drawing.Color.White;
			appearance229.BackColorDisabled2 = System.Drawing.Color.White;
			appearance229.BorderColor = System.Drawing.Color.Black;
			appearance229.BorderColor3DBase = System.Drawing.Color.Black;
			appearance229.ForeColor = System.Drawing.Color.Black;
			appearance229.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorNoviembre.Appearance = appearance229;
			this.txtProyeccionMargenDistribuidorNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorNoviembre.Enabled = false;
			this.txtProyeccionMargenDistribuidorNoviembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorNoviembre.Location = new System.Drawing.Point(938, 232);
			this.txtProyeccionMargenDistribuidorNoviembre.Name = "txtProyeccionMargenDistribuidorNoviembre";
			this.txtProyeccionMargenDistribuidorNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorNoviembre.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorNoviembre.TabIndex = 1219;
			// 
			// txtProyeccionMargenDistribuidorDiciembre
			// 
			appearance230.BackColorDisabled = System.Drawing.Color.White;
			appearance230.BackColorDisabled2 = System.Drawing.Color.White;
			appearance230.BorderColor = System.Drawing.Color.Black;
			appearance230.BorderColor3DBase = System.Drawing.Color.Black;
			appearance230.ForeColor = System.Drawing.Color.Black;
			appearance230.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorDiciembre.Appearance = appearance230;
			this.txtProyeccionMargenDistribuidorDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorDiciembre.Enabled = false;
			this.txtProyeccionMargenDistribuidorDiciembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorDiciembre.Location = new System.Drawing.Point(1026, 232);
			this.txtProyeccionMargenDistribuidorDiciembre.Name = "txtProyeccionMargenDistribuidorDiciembre";
			this.txtProyeccionMargenDistribuidorDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorDiciembre.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorDiciembre.TabIndex = 1218;
			// 
			// txtProyeccionMargenDistribuidorSeptiembre
			// 
			appearance231.BackColorDisabled = System.Drawing.Color.White;
			appearance231.BackColorDisabled2 = System.Drawing.Color.White;
			appearance231.BorderColor = System.Drawing.Color.Black;
			appearance231.BorderColor3DBase = System.Drawing.Color.Black;
			appearance231.ForeColor = System.Drawing.Color.Black;
			appearance231.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorSeptiembre.Appearance = appearance231;
			this.txtProyeccionMargenDistribuidorSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorSeptiembre.Enabled = false;
			this.txtProyeccionMargenDistribuidorSeptiembre.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorSeptiembre.Location = new System.Drawing.Point(762, 232);
			this.txtProyeccionMargenDistribuidorSeptiembre.Name = "txtProyeccionMargenDistribuidorSeptiembre";
			this.txtProyeccionMargenDistribuidorSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorSeptiembre.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorSeptiembre.TabIndex = 1217;
			// 
			// txtProyeccionMargenDistribuidorOctubre
			// 
			appearance232.BackColorDisabled = System.Drawing.Color.White;
			appearance232.BackColorDisabled2 = System.Drawing.Color.White;
			appearance232.BorderColor = System.Drawing.Color.Black;
			appearance232.BorderColor3DBase = System.Drawing.Color.Black;
			appearance232.ForeColor = System.Drawing.Color.Black;
			appearance232.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorOctubre.Appearance = appearance232;
			this.txtProyeccionMargenDistribuidorOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorOctubre.Enabled = false;
			this.txtProyeccionMargenDistribuidorOctubre.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorOctubre.Location = new System.Drawing.Point(850, 232);
			this.txtProyeccionMargenDistribuidorOctubre.Name = "txtProyeccionMargenDistribuidorOctubre";
			this.txtProyeccionMargenDistribuidorOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorOctubre.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorOctubre.TabIndex = 1216;
			// 
			// txtProyeccionMargenDistribuidorJulio
			// 
			appearance233.BackColorDisabled = System.Drawing.Color.White;
			appearance233.BackColorDisabled2 = System.Drawing.Color.White;
			appearance233.BorderColor = System.Drawing.Color.Black;
			appearance233.BorderColor3DBase = System.Drawing.Color.Black;
			appearance233.ForeColor = System.Drawing.Color.Black;
			appearance233.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorJulio.Appearance = appearance233;
			this.txtProyeccionMargenDistribuidorJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorJulio.Enabled = false;
			this.txtProyeccionMargenDistribuidorJulio.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorJulio.Location = new System.Drawing.Point(586, 232);
			this.txtProyeccionMargenDistribuidorJulio.Name = "txtProyeccionMargenDistribuidorJulio";
			this.txtProyeccionMargenDistribuidorJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorJulio.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorJulio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorJulio.TabIndex = 1215;
			// 
			// txtProyeccionMargenDistribuidorAgosto
			// 
			appearance234.BackColorDisabled = System.Drawing.Color.White;
			appearance234.BackColorDisabled2 = System.Drawing.Color.White;
			appearance234.BorderColor = System.Drawing.Color.Black;
			appearance234.BorderColor3DBase = System.Drawing.Color.Black;
			appearance234.ForeColor = System.Drawing.Color.Black;
			appearance234.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorAgosto.Appearance = appearance234;
			this.txtProyeccionMargenDistribuidorAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorAgosto.Enabled = false;
			this.txtProyeccionMargenDistribuidorAgosto.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorAgosto.Location = new System.Drawing.Point(674, 232);
			this.txtProyeccionMargenDistribuidorAgosto.Name = "txtProyeccionMargenDistribuidorAgosto";
			this.txtProyeccionMargenDistribuidorAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorAgosto.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorAgosto.TabIndex = 1214;
			// 
			// txtProyeccionMargenDistribuidorMayo
			// 
			appearance235.BackColorDisabled = System.Drawing.Color.White;
			appearance235.BackColorDisabled2 = System.Drawing.Color.White;
			appearance235.BorderColor = System.Drawing.Color.Black;
			appearance235.BorderColor3DBase = System.Drawing.Color.Black;
			appearance235.ForeColor = System.Drawing.Color.Black;
			appearance235.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorMayo.Appearance = appearance235;
			this.txtProyeccionMargenDistribuidorMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorMayo.Enabled = false;
			this.txtProyeccionMargenDistribuidorMayo.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorMayo.Location = new System.Drawing.Point(410, 232);
			this.txtProyeccionMargenDistribuidorMayo.Name = "txtProyeccionMargenDistribuidorMayo";
			this.txtProyeccionMargenDistribuidorMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorMayo.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorMayo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorMayo.TabIndex = 1213;
			// 
			// txtProyeccionMargenDistribuidorJunio
			// 
			appearance236.BackColorDisabled = System.Drawing.Color.White;
			appearance236.BackColorDisabled2 = System.Drawing.Color.White;
			appearance236.BorderColor = System.Drawing.Color.Black;
			appearance236.BorderColor3DBase = System.Drawing.Color.Black;
			appearance236.ForeColor = System.Drawing.Color.Black;
			appearance236.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorJunio.Appearance = appearance236;
			this.txtProyeccionMargenDistribuidorJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorJunio.Enabled = false;
			this.txtProyeccionMargenDistribuidorJunio.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorJunio.Location = new System.Drawing.Point(498, 232);
			this.txtProyeccionMargenDistribuidorJunio.Name = "txtProyeccionMargenDistribuidorJunio";
			this.txtProyeccionMargenDistribuidorJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorJunio.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorJunio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorJunio.TabIndex = 1212;
			// 
			// txtProyeccionMargenDistribuidorMarzo
			// 
			appearance237.BackColorDisabled = System.Drawing.Color.White;
			appearance237.BackColorDisabled2 = System.Drawing.Color.White;
			appearance237.BorderColor = System.Drawing.Color.Black;
			appearance237.BorderColor3DBase = System.Drawing.Color.Black;
			appearance237.ForeColor = System.Drawing.Color.Black;
			appearance237.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorMarzo.Appearance = appearance237;
			this.txtProyeccionMargenDistribuidorMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorMarzo.Enabled = false;
			this.txtProyeccionMargenDistribuidorMarzo.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorMarzo.Location = new System.Drawing.Point(234, 232);
			this.txtProyeccionMargenDistribuidorMarzo.Name = "txtProyeccionMargenDistribuidorMarzo";
			this.txtProyeccionMargenDistribuidorMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorMarzo.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorMarzo.TabIndex = 1211;
			// 
			// txtProyeccionMargenDistribuidorAbril
			// 
			appearance238.BackColorDisabled = System.Drawing.Color.White;
			appearance238.BackColorDisabled2 = System.Drawing.Color.White;
			appearance238.BorderColor = System.Drawing.Color.Black;
			appearance238.BorderColor3DBase = System.Drawing.Color.Black;
			appearance238.ForeColor = System.Drawing.Color.Black;
			appearance238.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorAbril.Appearance = appearance238;
			this.txtProyeccionMargenDistribuidorAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorAbril.Enabled = false;
			this.txtProyeccionMargenDistribuidorAbril.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorAbril.Location = new System.Drawing.Point(322, 232);
			this.txtProyeccionMargenDistribuidorAbril.Name = "txtProyeccionMargenDistribuidorAbril";
			this.txtProyeccionMargenDistribuidorAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorAbril.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorAbril.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorAbril.TabIndex = 1210;
			// 
			// txtProyeccionMargenDistribuidorEnero
			// 
			appearance239.BackColorDisabled = System.Drawing.Color.White;
			appearance239.BackColorDisabled2 = System.Drawing.Color.White;
			appearance239.BorderColor = System.Drawing.Color.Black;
			appearance239.BorderColor3DBase = System.Drawing.Color.Black;
			appearance239.ForeColor = System.Drawing.Color.Black;
			appearance239.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorEnero.Appearance = appearance239;
			this.txtProyeccionMargenDistribuidorEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorEnero.Enabled = false;
			this.txtProyeccionMargenDistribuidorEnero.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorEnero.Location = new System.Drawing.Point(58, 232);
			this.txtProyeccionMargenDistribuidorEnero.Name = "txtProyeccionMargenDistribuidorEnero";
			this.txtProyeccionMargenDistribuidorEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorEnero.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorEnero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorEnero.TabIndex = 1209;
			// 
			// txtProyeccionMargenDistribuidorFebrero
			// 
			appearance240.BackColorDisabled = System.Drawing.Color.White;
			appearance240.BackColorDisabled2 = System.Drawing.Color.White;
			appearance240.BorderColor = System.Drawing.Color.Black;
			appearance240.BorderColor3DBase = System.Drawing.Color.Black;
			appearance240.ForeColor = System.Drawing.Color.Black;
			appearance240.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionMargenDistribuidorFebrero.Appearance = appearance240;
			this.txtProyeccionMargenDistribuidorFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionMargenDistribuidorFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMargenDistribuidorFebrero.Enabled = false;
			this.txtProyeccionMargenDistribuidorFebrero.FormatString = "#,##0.00";
			this.txtProyeccionMargenDistribuidorFebrero.Location = new System.Drawing.Point(146, 232);
			this.txtProyeccionMargenDistribuidorFebrero.Name = "txtProyeccionMargenDistribuidorFebrero";
			this.txtProyeccionMargenDistribuidorFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMargenDistribuidorFebrero.PromptChar = ' ';
			this.txtProyeccionMargenDistribuidorFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionMargenDistribuidorFebrero.TabIndex = 1208;
			// 
			// txtProyeccionVentaCredito
			// 
			appearance241.BackColorDisabled = System.Drawing.Color.White;
			appearance241.BackColorDisabled2 = System.Drawing.Color.White;
			appearance241.BorderColor = System.Drawing.Color.Black;
			appearance241.BorderColor3DBase = System.Drawing.Color.Black;
			appearance241.FontData.BoldAsString = "True";
			appearance241.ForeColor = System.Drawing.Color.Firebrick;
			appearance241.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaCredito.Appearance = appearance241;
			this.txtProyeccionVentaCredito.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCredito.Enabled = false;
			this.txtProyeccionVentaCredito.FormatString = "#,##0.00";
			this.txtProyeccionVentaCredito.Location = new System.Drawing.Point(1114, 168);
			this.txtProyeccionVentaCredito.Name = "txtProyeccionVentaCredito";
			this.txtProyeccionVentaCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCredito.PromptChar = ' ';
			this.txtProyeccionVentaCredito.Size = new System.Drawing.Size(100, 19);
			this.txtProyeccionVentaCredito.TabIndex = 1207;
			// 
			// txtProyeccionVentaCreditoNoviembre
			// 
			appearance242.BackColorDisabled = System.Drawing.Color.White;
			appearance242.BackColorDisabled2 = System.Drawing.Color.White;
			appearance242.BorderColor = System.Drawing.Color.Black;
			appearance242.BorderColor3DBase = System.Drawing.Color.Black;
			appearance242.ForeColor = System.Drawing.Color.Black;
			appearance242.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoNoviembre.Appearance = appearance242;
			this.txtProyeccionVentaCreditoNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoNoviembre.Enabled = false;
			this.txtProyeccionVentaCreditoNoviembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoNoviembre.Location = new System.Drawing.Point(938, 168);
			this.txtProyeccionVentaCreditoNoviembre.Name = "txtProyeccionVentaCreditoNoviembre";
			this.txtProyeccionVentaCreditoNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoNoviembre.PromptChar = ' ';
			this.txtProyeccionVentaCreditoNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoNoviembre.TabIndex = 1206;
			// 
			// txtProyeccionVentaCreditoDiciembre
			// 
			appearance243.BackColorDisabled = System.Drawing.Color.White;
			appearance243.BackColorDisabled2 = System.Drawing.Color.White;
			appearance243.BorderColor = System.Drawing.Color.Black;
			appearance243.BorderColor3DBase = System.Drawing.Color.Black;
			appearance243.ForeColor = System.Drawing.Color.Black;
			appearance243.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoDiciembre.Appearance = appearance243;
			this.txtProyeccionVentaCreditoDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoDiciembre.Enabled = false;
			this.txtProyeccionVentaCreditoDiciembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoDiciembre.Location = new System.Drawing.Point(1026, 168);
			this.txtProyeccionVentaCreditoDiciembre.Name = "txtProyeccionVentaCreditoDiciembre";
			this.txtProyeccionVentaCreditoDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoDiciembre.PromptChar = ' ';
			this.txtProyeccionVentaCreditoDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoDiciembre.TabIndex = 1205;
			// 
			// txtProyeccionVentaCreditoSeptiembre
			// 
			appearance244.BackColorDisabled = System.Drawing.Color.White;
			appearance244.BackColorDisabled2 = System.Drawing.Color.White;
			appearance244.BorderColor = System.Drawing.Color.Black;
			appearance244.BorderColor3DBase = System.Drawing.Color.Black;
			appearance244.ForeColor = System.Drawing.Color.Black;
			appearance244.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoSeptiembre.Appearance = appearance244;
			this.txtProyeccionVentaCreditoSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoSeptiembre.Enabled = false;
			this.txtProyeccionVentaCreditoSeptiembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoSeptiembre.Location = new System.Drawing.Point(762, 168);
			this.txtProyeccionVentaCreditoSeptiembre.Name = "txtProyeccionVentaCreditoSeptiembre";
			this.txtProyeccionVentaCreditoSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoSeptiembre.PromptChar = ' ';
			this.txtProyeccionVentaCreditoSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoSeptiembre.TabIndex = 1204;
			// 
			// txtProyeccionVentaCreditoOctubre
			// 
			appearance245.BackColorDisabled = System.Drawing.Color.White;
			appearance245.BackColorDisabled2 = System.Drawing.Color.White;
			appearance245.BorderColor = System.Drawing.Color.Black;
			appearance245.BorderColor3DBase = System.Drawing.Color.Black;
			appearance245.ForeColor = System.Drawing.Color.Black;
			appearance245.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoOctubre.Appearance = appearance245;
			this.txtProyeccionVentaCreditoOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoOctubre.Enabled = false;
			this.txtProyeccionVentaCreditoOctubre.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoOctubre.Location = new System.Drawing.Point(850, 168);
			this.txtProyeccionVentaCreditoOctubre.Name = "txtProyeccionVentaCreditoOctubre";
			this.txtProyeccionVentaCreditoOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoOctubre.PromptChar = ' ';
			this.txtProyeccionVentaCreditoOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoOctubre.TabIndex = 1203;
			// 
			// txtProyeccionVentaCreditoJulio
			// 
			appearance246.BackColorDisabled = System.Drawing.Color.White;
			appearance246.BackColorDisabled2 = System.Drawing.Color.White;
			appearance246.BorderColor = System.Drawing.Color.Black;
			appearance246.BorderColor3DBase = System.Drawing.Color.Black;
			appearance246.ForeColor = System.Drawing.Color.Black;
			appearance246.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoJulio.Appearance = appearance246;
			this.txtProyeccionVentaCreditoJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoJulio.Enabled = false;
			this.txtProyeccionVentaCreditoJulio.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoJulio.Location = new System.Drawing.Point(586, 168);
			this.txtProyeccionVentaCreditoJulio.Name = "txtProyeccionVentaCreditoJulio";
			this.txtProyeccionVentaCreditoJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoJulio.PromptChar = ' ';
			this.txtProyeccionVentaCreditoJulio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoJulio.TabIndex = 1202;
			// 
			// txtProyeccionVentaCreditoAgosto
			// 
			appearance247.BackColorDisabled = System.Drawing.Color.White;
			appearance247.BackColorDisabled2 = System.Drawing.Color.White;
			appearance247.BorderColor = System.Drawing.Color.Black;
			appearance247.BorderColor3DBase = System.Drawing.Color.Black;
			appearance247.ForeColor = System.Drawing.Color.Black;
			appearance247.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoAgosto.Appearance = appearance247;
			this.txtProyeccionVentaCreditoAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoAgosto.Enabled = false;
			this.txtProyeccionVentaCreditoAgosto.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoAgosto.Location = new System.Drawing.Point(674, 168);
			this.txtProyeccionVentaCreditoAgosto.Name = "txtProyeccionVentaCreditoAgosto";
			this.txtProyeccionVentaCreditoAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoAgosto.PromptChar = ' ';
			this.txtProyeccionVentaCreditoAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoAgosto.TabIndex = 1201;
			// 
			// txtProyeccionVentaCreditoMayo
			// 
			appearance248.BackColorDisabled = System.Drawing.Color.White;
			appearance248.BackColorDisabled2 = System.Drawing.Color.White;
			appearance248.BorderColor = System.Drawing.Color.Black;
			appearance248.BorderColor3DBase = System.Drawing.Color.Black;
			appearance248.ForeColor = System.Drawing.Color.Black;
			appearance248.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoMayo.Appearance = appearance248;
			this.txtProyeccionVentaCreditoMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoMayo.Enabled = false;
			this.txtProyeccionVentaCreditoMayo.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoMayo.Location = new System.Drawing.Point(410, 168);
			this.txtProyeccionVentaCreditoMayo.Name = "txtProyeccionVentaCreditoMayo";
			this.txtProyeccionVentaCreditoMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoMayo.PromptChar = ' ';
			this.txtProyeccionVentaCreditoMayo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoMayo.TabIndex = 1200;
			// 
			// txtProyeccionVentaCreditoJunio
			// 
			appearance249.BackColorDisabled = System.Drawing.Color.White;
			appearance249.BackColorDisabled2 = System.Drawing.Color.White;
			appearance249.BorderColor = System.Drawing.Color.Black;
			appearance249.BorderColor3DBase = System.Drawing.Color.Black;
			appearance249.ForeColor = System.Drawing.Color.Black;
			appearance249.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoJunio.Appearance = appearance249;
			this.txtProyeccionVentaCreditoJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoJunio.Enabled = false;
			this.txtProyeccionVentaCreditoJunio.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoJunio.Location = new System.Drawing.Point(498, 168);
			this.txtProyeccionVentaCreditoJunio.Name = "txtProyeccionVentaCreditoJunio";
			this.txtProyeccionVentaCreditoJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoJunio.PromptChar = ' ';
			this.txtProyeccionVentaCreditoJunio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoJunio.TabIndex = 1199;
			// 
			// txtProyeccionVentaCreditoMarzo
			// 
			appearance250.BackColorDisabled = System.Drawing.Color.White;
			appearance250.BackColorDisabled2 = System.Drawing.Color.White;
			appearance250.BorderColor = System.Drawing.Color.Black;
			appearance250.BorderColor3DBase = System.Drawing.Color.Black;
			appearance250.ForeColor = System.Drawing.Color.Black;
			appearance250.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoMarzo.Appearance = appearance250;
			this.txtProyeccionVentaCreditoMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoMarzo.Enabled = false;
			this.txtProyeccionVentaCreditoMarzo.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoMarzo.Location = new System.Drawing.Point(234, 168);
			this.txtProyeccionVentaCreditoMarzo.Name = "txtProyeccionVentaCreditoMarzo";
			this.txtProyeccionVentaCreditoMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoMarzo.PromptChar = ' ';
			this.txtProyeccionVentaCreditoMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoMarzo.TabIndex = 1198;
			// 
			// txtProyeccionVentaCreditoAbril
			// 
			appearance251.BackColorDisabled = System.Drawing.Color.White;
			appearance251.BackColorDisabled2 = System.Drawing.Color.White;
			appearance251.BorderColor = System.Drawing.Color.Black;
			appearance251.BorderColor3DBase = System.Drawing.Color.Black;
			appearance251.ForeColor = System.Drawing.Color.Black;
			appearance251.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoAbril.Appearance = appearance251;
			this.txtProyeccionVentaCreditoAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoAbril.Enabled = false;
			this.txtProyeccionVentaCreditoAbril.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoAbril.Location = new System.Drawing.Point(322, 168);
			this.txtProyeccionVentaCreditoAbril.Name = "txtProyeccionVentaCreditoAbril";
			this.txtProyeccionVentaCreditoAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoAbril.PromptChar = ' ';
			this.txtProyeccionVentaCreditoAbril.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoAbril.TabIndex = 1197;
			// 
			// txtProyeccionVentaCreditoEnero
			// 
			appearance252.BackColorDisabled = System.Drawing.Color.White;
			appearance252.BackColorDisabled2 = System.Drawing.Color.White;
			appearance252.BorderColor = System.Drawing.Color.Black;
			appearance252.BorderColor3DBase = System.Drawing.Color.Black;
			appearance252.ForeColor = System.Drawing.Color.Black;
			appearance252.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoEnero.Appearance = appearance252;
			this.txtProyeccionVentaCreditoEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoEnero.Enabled = false;
			this.txtProyeccionVentaCreditoEnero.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoEnero.Location = new System.Drawing.Point(58, 168);
			this.txtProyeccionVentaCreditoEnero.Name = "txtProyeccionVentaCreditoEnero";
			this.txtProyeccionVentaCreditoEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoEnero.PromptChar = ' ';
			this.txtProyeccionVentaCreditoEnero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoEnero.TabIndex = 1196;
			// 
			// txtProyeccionVentaCreditoFebrero
			// 
			appearance253.BackColorDisabled = System.Drawing.Color.White;
			appearance253.BackColorDisabled2 = System.Drawing.Color.White;
			appearance253.BorderColor = System.Drawing.Color.Black;
			appearance253.BorderColor3DBase = System.Drawing.Color.Black;
			appearance253.ForeColor = System.Drawing.Color.Black;
			appearance253.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaCreditoFebrero.Appearance = appearance253;
			this.txtProyeccionVentaCreditoFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaCreditoFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaCreditoFebrero.Enabled = false;
			this.txtProyeccionVentaCreditoFebrero.FormatString = "#,##0.00";
			this.txtProyeccionVentaCreditoFebrero.Location = new System.Drawing.Point(146, 168);
			this.txtProyeccionVentaCreditoFebrero.Name = "txtProyeccionVentaCreditoFebrero";
			this.txtProyeccionVentaCreditoFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaCreditoFebrero.PromptChar = ' ';
			this.txtProyeccionVentaCreditoFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaCreditoFebrero.TabIndex = 1195;
			// 
			// txtProyeccionVentaTarjeta
			// 
			appearance254.BackColorDisabled = System.Drawing.Color.White;
			appearance254.BackColorDisabled2 = System.Drawing.Color.White;
			appearance254.BorderColor = System.Drawing.Color.Black;
			appearance254.BorderColor3DBase = System.Drawing.Color.Black;
			appearance254.FontData.BoldAsString = "True";
			appearance254.ForeColor = System.Drawing.Color.Firebrick;
			appearance254.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaTarjeta.Appearance = appearance254;
			this.txtProyeccionVentaTarjeta.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjeta.Enabled = false;
			this.txtProyeccionVentaTarjeta.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjeta.Location = new System.Drawing.Point(1114, 144);
			this.txtProyeccionVentaTarjeta.Name = "txtProyeccionVentaTarjeta";
			this.txtProyeccionVentaTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjeta.PromptChar = ' ';
			this.txtProyeccionVentaTarjeta.Size = new System.Drawing.Size(100, 19);
			this.txtProyeccionVentaTarjeta.TabIndex = 1194;
			// 
			// txtProyeccionVentaTarjetaNoviembre
			// 
			appearance255.BackColorDisabled = System.Drawing.Color.White;
			appearance255.BackColorDisabled2 = System.Drawing.Color.White;
			appearance255.BorderColor = System.Drawing.Color.Black;
			appearance255.BorderColor3DBase = System.Drawing.Color.Black;
			appearance255.ForeColor = System.Drawing.Color.Black;
			appearance255.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaNoviembre.Appearance = appearance255;
			this.txtProyeccionVentaTarjetaNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaNoviembre.Enabled = false;
			this.txtProyeccionVentaTarjetaNoviembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaNoviembre.Location = new System.Drawing.Point(938, 144);
			this.txtProyeccionVentaTarjetaNoviembre.Name = "txtProyeccionVentaTarjetaNoviembre";
			this.txtProyeccionVentaTarjetaNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaNoviembre.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaNoviembre.TabIndex = 1193;
			// 
			// txtProyeccionVentaTarjetaDiciembre
			// 
			appearance256.BackColorDisabled = System.Drawing.Color.White;
			appearance256.BackColorDisabled2 = System.Drawing.Color.White;
			appearance256.BorderColor = System.Drawing.Color.Black;
			appearance256.BorderColor3DBase = System.Drawing.Color.Black;
			appearance256.ForeColor = System.Drawing.Color.Black;
			appearance256.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaDiciembre.Appearance = appearance256;
			this.txtProyeccionVentaTarjetaDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaDiciembre.Enabled = false;
			this.txtProyeccionVentaTarjetaDiciembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaDiciembre.Location = new System.Drawing.Point(1026, 144);
			this.txtProyeccionVentaTarjetaDiciembre.Name = "txtProyeccionVentaTarjetaDiciembre";
			this.txtProyeccionVentaTarjetaDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaDiciembre.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaDiciembre.TabIndex = 1192;
			// 
			// txtProyeccionVentaTarjetaSeptiembre
			// 
			appearance257.BackColorDisabled = System.Drawing.Color.White;
			appearance257.BackColorDisabled2 = System.Drawing.Color.White;
			appearance257.BorderColor = System.Drawing.Color.Black;
			appearance257.BorderColor3DBase = System.Drawing.Color.Black;
			appearance257.ForeColor = System.Drawing.Color.Black;
			appearance257.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaSeptiembre.Appearance = appearance257;
			this.txtProyeccionVentaTarjetaSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaSeptiembre.Enabled = false;
			this.txtProyeccionVentaTarjetaSeptiembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaSeptiembre.Location = new System.Drawing.Point(762, 144);
			this.txtProyeccionVentaTarjetaSeptiembre.Name = "txtProyeccionVentaTarjetaSeptiembre";
			this.txtProyeccionVentaTarjetaSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaSeptiembre.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaSeptiembre.TabIndex = 1191;
			// 
			// txtProyeccionVentaTarjetaOctubre
			// 
			appearance258.BackColorDisabled = System.Drawing.Color.White;
			appearance258.BackColorDisabled2 = System.Drawing.Color.White;
			appearance258.BorderColor = System.Drawing.Color.Black;
			appearance258.BorderColor3DBase = System.Drawing.Color.Black;
			appearance258.ForeColor = System.Drawing.Color.Black;
			appearance258.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaOctubre.Appearance = appearance258;
			this.txtProyeccionVentaTarjetaOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaOctubre.Enabled = false;
			this.txtProyeccionVentaTarjetaOctubre.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaOctubre.Location = new System.Drawing.Point(850, 144);
			this.txtProyeccionVentaTarjetaOctubre.Name = "txtProyeccionVentaTarjetaOctubre";
			this.txtProyeccionVentaTarjetaOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaOctubre.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaOctubre.TabIndex = 1190;
			// 
			// txtProyeccionVentaTarjetaJulio
			// 
			appearance259.BackColorDisabled = System.Drawing.Color.White;
			appearance259.BackColorDisabled2 = System.Drawing.Color.White;
			appearance259.BorderColor = System.Drawing.Color.Black;
			appearance259.BorderColor3DBase = System.Drawing.Color.Black;
			appearance259.ForeColor = System.Drawing.Color.Black;
			appearance259.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaJulio.Appearance = appearance259;
			this.txtProyeccionVentaTarjetaJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaJulio.Enabled = false;
			this.txtProyeccionVentaTarjetaJulio.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaJulio.Location = new System.Drawing.Point(586, 144);
			this.txtProyeccionVentaTarjetaJulio.Name = "txtProyeccionVentaTarjetaJulio";
			this.txtProyeccionVentaTarjetaJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaJulio.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaJulio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaJulio.TabIndex = 1189;
			// 
			// txtProyeccionVentaTarjetaAgosto
			// 
			appearance260.BackColorDisabled = System.Drawing.Color.White;
			appearance260.BackColorDisabled2 = System.Drawing.Color.White;
			appearance260.BorderColor = System.Drawing.Color.Black;
			appearance260.BorderColor3DBase = System.Drawing.Color.Black;
			appearance260.ForeColor = System.Drawing.Color.Black;
			appearance260.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaAgosto.Appearance = appearance260;
			this.txtProyeccionVentaTarjetaAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaAgosto.Enabled = false;
			this.txtProyeccionVentaTarjetaAgosto.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaAgosto.Location = new System.Drawing.Point(674, 144);
			this.txtProyeccionVentaTarjetaAgosto.Name = "txtProyeccionVentaTarjetaAgosto";
			this.txtProyeccionVentaTarjetaAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaAgosto.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaAgosto.TabIndex = 1188;
			// 
			// txtProyeccionVentaTarjetaMayo
			// 
			appearance261.BackColorDisabled = System.Drawing.Color.White;
			appearance261.BackColorDisabled2 = System.Drawing.Color.White;
			appearance261.BorderColor = System.Drawing.Color.Black;
			appearance261.BorderColor3DBase = System.Drawing.Color.Black;
			appearance261.ForeColor = System.Drawing.Color.Black;
			appearance261.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaMayo.Appearance = appearance261;
			this.txtProyeccionVentaTarjetaMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaMayo.Enabled = false;
			this.txtProyeccionVentaTarjetaMayo.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaMayo.Location = new System.Drawing.Point(410, 144);
			this.txtProyeccionVentaTarjetaMayo.Name = "txtProyeccionVentaTarjetaMayo";
			this.txtProyeccionVentaTarjetaMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaMayo.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaMayo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaMayo.TabIndex = 1187;
			// 
			// txtProyeccionVentaTarjetaJunio
			// 
			appearance262.BackColorDisabled = System.Drawing.Color.White;
			appearance262.BackColorDisabled2 = System.Drawing.Color.White;
			appearance262.BorderColor = System.Drawing.Color.Black;
			appearance262.BorderColor3DBase = System.Drawing.Color.Black;
			appearance262.ForeColor = System.Drawing.Color.Black;
			appearance262.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaJunio.Appearance = appearance262;
			this.txtProyeccionVentaTarjetaJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaJunio.Enabled = false;
			this.txtProyeccionVentaTarjetaJunio.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaJunio.Location = new System.Drawing.Point(498, 144);
			this.txtProyeccionVentaTarjetaJunio.Name = "txtProyeccionVentaTarjetaJunio";
			this.txtProyeccionVentaTarjetaJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaJunio.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaJunio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaJunio.TabIndex = 1186;
			// 
			// txtProyeccionVentaTarjetaMarzo
			// 
			appearance263.BackColorDisabled = System.Drawing.Color.White;
			appearance263.BackColorDisabled2 = System.Drawing.Color.White;
			appearance263.BorderColor = System.Drawing.Color.Black;
			appearance263.BorderColor3DBase = System.Drawing.Color.Black;
			appearance263.ForeColor = System.Drawing.Color.Black;
			appearance263.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaMarzo.Appearance = appearance263;
			this.txtProyeccionVentaTarjetaMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaMarzo.Enabled = false;
			this.txtProyeccionVentaTarjetaMarzo.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaMarzo.Location = new System.Drawing.Point(234, 144);
			this.txtProyeccionVentaTarjetaMarzo.Name = "txtProyeccionVentaTarjetaMarzo";
			this.txtProyeccionVentaTarjetaMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaMarzo.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaMarzo.TabIndex = 1185;
			// 
			// txtProyeccionVentaTarjetaAbril
			// 
			appearance264.BackColorDisabled = System.Drawing.Color.White;
			appearance264.BackColorDisabled2 = System.Drawing.Color.White;
			appearance264.BorderColor = System.Drawing.Color.Black;
			appearance264.BorderColor3DBase = System.Drawing.Color.Black;
			appearance264.ForeColor = System.Drawing.Color.Black;
			appearance264.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaAbril.Appearance = appearance264;
			this.txtProyeccionVentaTarjetaAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaAbril.Enabled = false;
			this.txtProyeccionVentaTarjetaAbril.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaAbril.Location = new System.Drawing.Point(322, 144);
			this.txtProyeccionVentaTarjetaAbril.Name = "txtProyeccionVentaTarjetaAbril";
			this.txtProyeccionVentaTarjetaAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaAbril.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaAbril.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaAbril.TabIndex = 1184;
			// 
			// txtProyeccionVentaTarjetaEnero
			// 
			appearance265.BackColorDisabled = System.Drawing.Color.White;
			appearance265.BackColorDisabled2 = System.Drawing.Color.White;
			appearance265.BorderColor = System.Drawing.Color.Black;
			appearance265.BorderColor3DBase = System.Drawing.Color.Black;
			appearance265.ForeColor = System.Drawing.Color.Black;
			appearance265.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaEnero.Appearance = appearance265;
			this.txtProyeccionVentaTarjetaEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaEnero.Enabled = false;
			this.txtProyeccionVentaTarjetaEnero.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaEnero.Location = new System.Drawing.Point(58, 144);
			this.txtProyeccionVentaTarjetaEnero.Name = "txtProyeccionVentaTarjetaEnero";
			this.txtProyeccionVentaTarjetaEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaEnero.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaEnero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaEnero.TabIndex = 1183;
			// 
			// txtProyeccionVentaTarjetaFebrero
			// 
			appearance266.BackColorDisabled = System.Drawing.Color.White;
			appearance266.BackColorDisabled2 = System.Drawing.Color.White;
			appearance266.BorderColor = System.Drawing.Color.Black;
			appearance266.BorderColor3DBase = System.Drawing.Color.Black;
			appearance266.ForeColor = System.Drawing.Color.Black;
			appearance266.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaTarjetaFebrero.Appearance = appearance266;
			this.txtProyeccionVentaTarjetaFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaTarjetaFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaTarjetaFebrero.Enabled = false;
			this.txtProyeccionVentaTarjetaFebrero.FormatString = "#,##0.00";
			this.txtProyeccionVentaTarjetaFebrero.Location = new System.Drawing.Point(146, 144);
			this.txtProyeccionVentaTarjetaFebrero.Name = "txtProyeccionVentaTarjetaFebrero";
			this.txtProyeccionVentaTarjetaFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaTarjetaFebrero.PromptChar = ' ';
			this.txtProyeccionVentaTarjetaFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaTarjetaFebrero.TabIndex = 1182;
			// 
			// txtProyeccionVentaEfectivo
			// 
			appearance267.BackColorDisabled = System.Drawing.Color.White;
			appearance267.BackColorDisabled2 = System.Drawing.Color.White;
			appearance267.BorderColor = System.Drawing.Color.Black;
			appearance267.BorderColor3DBase = System.Drawing.Color.Black;
			appearance267.FontData.BoldAsString = "True";
			appearance267.ForeColor = System.Drawing.Color.Firebrick;
			appearance267.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaEfectivo.Appearance = appearance267;
			this.txtProyeccionVentaEfectivo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivo.Enabled = false;
			this.txtProyeccionVentaEfectivo.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivo.Location = new System.Drawing.Point(1114, 120);
			this.txtProyeccionVentaEfectivo.Name = "txtProyeccionVentaEfectivo";
			this.txtProyeccionVentaEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivo.PromptChar = ' ';
			this.txtProyeccionVentaEfectivo.Size = new System.Drawing.Size(100, 19);
			this.txtProyeccionVentaEfectivo.TabIndex = 1181;
			// 
			// txtProyeccionVentaEfectivoNoviembre
			// 
			appearance268.BackColorDisabled = System.Drawing.Color.White;
			appearance268.BackColorDisabled2 = System.Drawing.Color.White;
			appearance268.BorderColor = System.Drawing.Color.Black;
			appearance268.BorderColor3DBase = System.Drawing.Color.Black;
			appearance268.ForeColor = System.Drawing.Color.Black;
			appearance268.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoNoviembre.Appearance = appearance268;
			this.txtProyeccionVentaEfectivoNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoNoviembre.Enabled = false;
			this.txtProyeccionVentaEfectivoNoviembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoNoviembre.Location = new System.Drawing.Point(938, 120);
			this.txtProyeccionVentaEfectivoNoviembre.Name = "txtProyeccionVentaEfectivoNoviembre";
			this.txtProyeccionVentaEfectivoNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoNoviembre.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoNoviembre.TabIndex = 1180;
			// 
			// txtProyeccionVentaEfectivoDiciembre
			// 
			appearance269.BackColorDisabled = System.Drawing.Color.White;
			appearance269.BackColorDisabled2 = System.Drawing.Color.White;
			appearance269.BorderColor = System.Drawing.Color.Black;
			appearance269.BorderColor3DBase = System.Drawing.Color.Black;
			appearance269.ForeColor = System.Drawing.Color.Black;
			appearance269.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoDiciembre.Appearance = appearance269;
			this.txtProyeccionVentaEfectivoDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoDiciembre.Enabled = false;
			this.txtProyeccionVentaEfectivoDiciembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoDiciembre.Location = new System.Drawing.Point(1026, 120);
			this.txtProyeccionVentaEfectivoDiciembre.Name = "txtProyeccionVentaEfectivoDiciembre";
			this.txtProyeccionVentaEfectivoDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoDiciembre.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoDiciembre.TabIndex = 1179;
			// 
			// txtProyeccionVentaEfectivoSeptiembre
			// 
			appearance270.BackColorDisabled = System.Drawing.Color.White;
			appearance270.BackColorDisabled2 = System.Drawing.Color.White;
			appearance270.BorderColor = System.Drawing.Color.Black;
			appearance270.BorderColor3DBase = System.Drawing.Color.Black;
			appearance270.ForeColor = System.Drawing.Color.Black;
			appearance270.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoSeptiembre.Appearance = appearance270;
			this.txtProyeccionVentaEfectivoSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoSeptiembre.Enabled = false;
			this.txtProyeccionVentaEfectivoSeptiembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoSeptiembre.Location = new System.Drawing.Point(762, 120);
			this.txtProyeccionVentaEfectivoSeptiembre.Name = "txtProyeccionVentaEfectivoSeptiembre";
			this.txtProyeccionVentaEfectivoSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoSeptiembre.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoSeptiembre.TabIndex = 1178;
			// 
			// txtProyeccionVentaEfectivoOctubre
			// 
			appearance271.BackColorDisabled = System.Drawing.Color.White;
			appearance271.BackColorDisabled2 = System.Drawing.Color.White;
			appearance271.BorderColor = System.Drawing.Color.Black;
			appearance271.BorderColor3DBase = System.Drawing.Color.Black;
			appearance271.ForeColor = System.Drawing.Color.Black;
			appearance271.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoOctubre.Appearance = appearance271;
			this.txtProyeccionVentaEfectivoOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoOctubre.Enabled = false;
			this.txtProyeccionVentaEfectivoOctubre.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoOctubre.Location = new System.Drawing.Point(850, 120);
			this.txtProyeccionVentaEfectivoOctubre.Name = "txtProyeccionVentaEfectivoOctubre";
			this.txtProyeccionVentaEfectivoOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoOctubre.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoOctubre.TabIndex = 1177;
			// 
			// txtProyeccionVentaEfectivoJulio
			// 
			appearance272.BackColorDisabled = System.Drawing.Color.White;
			appearance272.BackColorDisabled2 = System.Drawing.Color.White;
			appearance272.BorderColor = System.Drawing.Color.Black;
			appearance272.BorderColor3DBase = System.Drawing.Color.Black;
			appearance272.ForeColor = System.Drawing.Color.Black;
			appearance272.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoJulio.Appearance = appearance272;
			this.txtProyeccionVentaEfectivoJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoJulio.Enabled = false;
			this.txtProyeccionVentaEfectivoJulio.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoJulio.Location = new System.Drawing.Point(586, 120);
			this.txtProyeccionVentaEfectivoJulio.Name = "txtProyeccionVentaEfectivoJulio";
			this.txtProyeccionVentaEfectivoJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoJulio.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoJulio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoJulio.TabIndex = 1176;
			// 
			// txtProyeccionVentaEfectivoAgosto
			// 
			appearance273.BackColorDisabled = System.Drawing.Color.White;
			appearance273.BackColorDisabled2 = System.Drawing.Color.White;
			appearance273.BorderColor = System.Drawing.Color.Black;
			appearance273.BorderColor3DBase = System.Drawing.Color.Black;
			appearance273.ForeColor = System.Drawing.Color.Black;
			appearance273.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoAgosto.Appearance = appearance273;
			this.txtProyeccionVentaEfectivoAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoAgosto.Enabled = false;
			this.txtProyeccionVentaEfectivoAgosto.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoAgosto.Location = new System.Drawing.Point(674, 120);
			this.txtProyeccionVentaEfectivoAgosto.Name = "txtProyeccionVentaEfectivoAgosto";
			this.txtProyeccionVentaEfectivoAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoAgosto.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoAgosto.TabIndex = 1175;
			// 
			// txtProyeccionVentaEfectivoMayo
			// 
			appearance274.BackColorDisabled = System.Drawing.Color.White;
			appearance274.BackColorDisabled2 = System.Drawing.Color.White;
			appearance274.BorderColor = System.Drawing.Color.Black;
			appearance274.BorderColor3DBase = System.Drawing.Color.Black;
			appearance274.ForeColor = System.Drawing.Color.Black;
			appearance274.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoMayo.Appearance = appearance274;
			this.txtProyeccionVentaEfectivoMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoMayo.Enabled = false;
			this.txtProyeccionVentaEfectivoMayo.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoMayo.Location = new System.Drawing.Point(410, 120);
			this.txtProyeccionVentaEfectivoMayo.Name = "txtProyeccionVentaEfectivoMayo";
			this.txtProyeccionVentaEfectivoMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoMayo.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoMayo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoMayo.TabIndex = 1174;
			// 
			// txtProyeccionVentaEfectivoJunio
			// 
			appearance275.BackColorDisabled = System.Drawing.Color.White;
			appearance275.BackColorDisabled2 = System.Drawing.Color.White;
			appearance275.BorderColor = System.Drawing.Color.Black;
			appearance275.BorderColor3DBase = System.Drawing.Color.Black;
			appearance275.ForeColor = System.Drawing.Color.Black;
			appearance275.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoJunio.Appearance = appearance275;
			this.txtProyeccionVentaEfectivoJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoJunio.Enabled = false;
			this.txtProyeccionVentaEfectivoJunio.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoJunio.Location = new System.Drawing.Point(498, 120);
			this.txtProyeccionVentaEfectivoJunio.Name = "txtProyeccionVentaEfectivoJunio";
			this.txtProyeccionVentaEfectivoJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoJunio.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoJunio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoJunio.TabIndex = 1173;
			// 
			// txtProyeccionVentaEfectivoMarzo
			// 
			appearance276.BackColorDisabled = System.Drawing.Color.White;
			appearance276.BackColorDisabled2 = System.Drawing.Color.White;
			appearance276.BorderColor = System.Drawing.Color.Black;
			appearance276.BorderColor3DBase = System.Drawing.Color.Black;
			appearance276.ForeColor = System.Drawing.Color.Black;
			appearance276.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoMarzo.Appearance = appearance276;
			this.txtProyeccionVentaEfectivoMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoMarzo.Enabled = false;
			this.txtProyeccionVentaEfectivoMarzo.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoMarzo.Location = new System.Drawing.Point(234, 120);
			this.txtProyeccionVentaEfectivoMarzo.Name = "txtProyeccionVentaEfectivoMarzo";
			this.txtProyeccionVentaEfectivoMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoMarzo.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoMarzo.TabIndex = 1172;
			// 
			// txtProyeccionVentaEfectivoAbril
			// 
			appearance277.BackColorDisabled = System.Drawing.Color.White;
			appearance277.BackColorDisabled2 = System.Drawing.Color.White;
			appearance277.BorderColor = System.Drawing.Color.Black;
			appearance277.BorderColor3DBase = System.Drawing.Color.Black;
			appearance277.ForeColor = System.Drawing.Color.Black;
			appearance277.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoAbril.Appearance = appearance277;
			this.txtProyeccionVentaEfectivoAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoAbril.Enabled = false;
			this.txtProyeccionVentaEfectivoAbril.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoAbril.Location = new System.Drawing.Point(322, 120);
			this.txtProyeccionVentaEfectivoAbril.Name = "txtProyeccionVentaEfectivoAbril";
			this.txtProyeccionVentaEfectivoAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoAbril.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoAbril.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoAbril.TabIndex = 1171;
			// 
			// txtProyeccionVentaEfectivoEnero
			// 
			appearance278.BackColorDisabled = System.Drawing.Color.White;
			appearance278.BackColorDisabled2 = System.Drawing.Color.White;
			appearance278.BorderColor = System.Drawing.Color.Black;
			appearance278.BorderColor3DBase = System.Drawing.Color.Black;
			appearance278.ForeColor = System.Drawing.Color.Black;
			appearance278.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoEnero.Appearance = appearance278;
			this.txtProyeccionVentaEfectivoEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoEnero.Enabled = false;
			this.txtProyeccionVentaEfectivoEnero.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoEnero.Location = new System.Drawing.Point(58, 120);
			this.txtProyeccionVentaEfectivoEnero.Name = "txtProyeccionVentaEfectivoEnero";
			this.txtProyeccionVentaEfectivoEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoEnero.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoEnero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoEnero.TabIndex = 1170;
			// 
			// txtProyeccionVentaEfectivoFebrero
			// 
			appearance279.BackColorDisabled = System.Drawing.Color.White;
			appearance279.BackColorDisabled2 = System.Drawing.Color.White;
			appearance279.BorderColor = System.Drawing.Color.Black;
			appearance279.BorderColor3DBase = System.Drawing.Color.Black;
			appearance279.ForeColor = System.Drawing.Color.Black;
			appearance279.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaEfectivoFebrero.Appearance = appearance279;
			this.txtProyeccionVentaEfectivoFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaEfectivoFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaEfectivoFebrero.Enabled = false;
			this.txtProyeccionVentaEfectivoFebrero.FormatString = "#,##0.00";
			this.txtProyeccionVentaEfectivoFebrero.Location = new System.Drawing.Point(146, 120);
			this.txtProyeccionVentaEfectivoFebrero.Name = "txtProyeccionVentaEfectivoFebrero";
			this.txtProyeccionVentaEfectivoFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaEfectivoFebrero.PromptChar = ' ';
			this.txtProyeccionVentaEfectivoFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaEfectivoFebrero.TabIndex = 1169;
			// 
			// label16
			// 
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.ForeColor = System.Drawing.Color.Firebrick;
			this.label16.Location = new System.Drawing.Point(1114, 40);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(100, 16);
			this.label16.TabIndex = 1168;
			this.label16.Text = "TOTAL";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label12
			// 
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.ForeColor = System.Drawing.Color.Firebrick;
			this.label12.Location = new System.Drawing.Point(1026, 40);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(80, 16);
			this.label12.TabIndex = 1167;
			this.label12.Text = "DICIEMBRE";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label13
			// 
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.ForeColor = System.Drawing.Color.Firebrick;
			this.label13.Location = new System.Drawing.Point(938, 40);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(80, 16);
			this.label13.TabIndex = 1166;
			this.label13.Text = "NOVIEMBRE";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label14
			// 
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.ForeColor = System.Drawing.Color.Firebrick;
			this.label14.Location = new System.Drawing.Point(850, 40);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(80, 16);
			this.label14.TabIndex = 1165;
			this.label14.Text = "OCTUBRE";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label15
			// 
			this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.ForeColor = System.Drawing.Color.Firebrick;
			this.label15.Location = new System.Drawing.Point(762, 40);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(80, 16);
			this.label15.TabIndex = 1164;
			this.label15.Text = "SEPTIEMBRE";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label8
			// 
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.Firebrick;
			this.label8.Location = new System.Drawing.Point(674, 40);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 1163;
			this.label8.Text = "AGOSTO";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label9
			// 
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.Firebrick;
			this.label9.Location = new System.Drawing.Point(586, 40);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 16);
			this.label9.TabIndex = 1162;
			this.label9.Text = "JULIO";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label10
			// 
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.ForeColor = System.Drawing.Color.Firebrick;
			this.label10.Location = new System.Drawing.Point(498, 40);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(80, 16);
			this.label10.TabIndex = 1161;
			this.label10.Text = "JUNIO";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label11
			// 
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.ForeColor = System.Drawing.Color.Firebrick;
			this.label11.Location = new System.Drawing.Point(410, 40);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(80, 16);
			this.label11.TabIndex = 1160;
			this.label11.Text = "MAYO";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label7
			// 
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.ForeColor = System.Drawing.Color.Firebrick;
			this.label7.Location = new System.Drawing.Point(322, 40);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 16);
			this.label7.TabIndex = 1159;
			this.label7.Text = "ABRIL";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label6
			// 
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.ForeColor = System.Drawing.Color.Firebrick;
			this.label6.Location = new System.Drawing.Point(234, 40);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 16);
			this.label6.TabIndex = 1158;
			this.label6.Text = "MARZO";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.ForeColor = System.Drawing.Color.Firebrick;
			this.label4.Location = new System.Drawing.Point(146, 40);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 16);
			this.label4.TabIndex = 1157;
			this.label4.Text = "FEBRERO";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Firebrick;
			this.label3.Location = new System.Drawing.Point(58, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(80, 16);
			this.label3.TabIndex = 1156;
			this.label3.Text = "ENERO";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtProyeccionVentaDistribuidor
			// 
			appearance280.BackColorDisabled = System.Drawing.Color.White;
			appearance280.BackColorDisabled2 = System.Drawing.Color.White;
			appearance280.BorderColor = System.Drawing.Color.Black;
			appearance280.BorderColor3DBase = System.Drawing.Color.Black;
			appearance280.FontData.BoldAsString = "True";
			appearance280.ForeColor = System.Drawing.Color.Firebrick;
			appearance280.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionVentaDistribuidor.Appearance = appearance280;
			this.txtProyeccionVentaDistribuidor.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidor.Enabled = false;
			this.txtProyeccionVentaDistribuidor.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidor.Location = new System.Drawing.Point(1114, 96);
			this.txtProyeccionVentaDistribuidor.Name = "txtProyeccionVentaDistribuidor";
			this.txtProyeccionVentaDistribuidor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidor.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidor.Size = new System.Drawing.Size(100, 19);
			this.txtProyeccionVentaDistribuidor.TabIndex = 1155;
			// 
			// txtIngresosTotal
			// 
			appearance281.FontData.BoldAsString = "True";
			appearance281.ForeColor = System.Drawing.Color.Firebrick;
			appearance281.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtIngresosTotal.Appearance = appearance281;
			this.txtIngresosTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosTotal.Enabled = false;
			this.txtIngresosTotal.FormatString = "#,##0.00";
			this.txtIngresosTotal.Location = new System.Drawing.Point(1114, 56);
			this.txtIngresosTotal.Name = "txtIngresosTotal";
			this.txtIngresosTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosTotal.PromptChar = ' ';
			this.txtIngresosTotal.Size = new System.Drawing.Size(100, 21);
			this.txtIngresosTotal.TabIndex = 1154;
			// 
			// txtProyeccionVentaDistribuidorNoviembre
			// 
			appearance282.BackColorDisabled = System.Drawing.Color.White;
			appearance282.BackColorDisabled2 = System.Drawing.Color.White;
			appearance282.BorderColor = System.Drawing.Color.Black;
			appearance282.BorderColor3DBase = System.Drawing.Color.Black;
			appearance282.ForeColor = System.Drawing.Color.Black;
			appearance282.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorNoviembre.Appearance = appearance282;
			this.txtProyeccionVentaDistribuidorNoviembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorNoviembre.Enabled = false;
			this.txtProyeccionVentaDistribuidorNoviembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorNoviembre.Location = new System.Drawing.Point(938, 96);
			this.txtProyeccionVentaDistribuidorNoviembre.Name = "txtProyeccionVentaDistribuidorNoviembre";
			this.txtProyeccionVentaDistribuidorNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorNoviembre.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorNoviembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorNoviembre.TabIndex = 1153;
			// 
			// txtProyeccionVentaDistribuidorDiciembre
			// 
			appearance283.BackColorDisabled = System.Drawing.Color.White;
			appearance283.BackColorDisabled2 = System.Drawing.Color.White;
			appearance283.BorderColor = System.Drawing.Color.Black;
			appearance283.BorderColor3DBase = System.Drawing.Color.Black;
			appearance283.ForeColor = System.Drawing.Color.Black;
			appearance283.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorDiciembre.Appearance = appearance283;
			this.txtProyeccionVentaDistribuidorDiciembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorDiciembre.Enabled = false;
			this.txtProyeccionVentaDistribuidorDiciembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorDiciembre.Location = new System.Drawing.Point(1026, 96);
			this.txtProyeccionVentaDistribuidorDiciembre.Name = "txtProyeccionVentaDistribuidorDiciembre";
			this.txtProyeccionVentaDistribuidorDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorDiciembre.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorDiciembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorDiciembre.TabIndex = 1152;
			// 
			// txtProyeccionDiciembreVenta
			// 
			appearance284.ForeColor = System.Drawing.Color.Firebrick;
			appearance284.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionDiciembreVenta.Appearance = appearance284;
			this.txtProyeccionDiciembreVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionDiciembreVenta.Enabled = false;
			this.txtProyeccionDiciembreVenta.FormatString = "#,##0.00";
			this.txtProyeccionDiciembreVenta.Location = new System.Drawing.Point(1026, 56);
			this.txtProyeccionDiciembreVenta.Name = "txtProyeccionDiciembreVenta";
			this.txtProyeccionDiciembreVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionDiciembreVenta.PromptChar = ' ';
			this.txtProyeccionDiciembreVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionDiciembreVenta.TabIndex = 1151;
			// 
			// txtProyeccionNoviembreVenta
			// 
			appearance285.ForeColor = System.Drawing.Color.Firebrick;
			appearance285.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionNoviembreVenta.Appearance = appearance285;
			this.txtProyeccionNoviembreVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionNoviembreVenta.Enabled = false;
			this.txtProyeccionNoviembreVenta.FormatString = "#,##0.00";
			this.txtProyeccionNoviembreVenta.Location = new System.Drawing.Point(938, 56);
			this.txtProyeccionNoviembreVenta.Name = "txtProyeccionNoviembreVenta";
			this.txtProyeccionNoviembreVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionNoviembreVenta.PromptChar = ' ';
			this.txtProyeccionNoviembreVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionNoviembreVenta.TabIndex = 1150;
			// 
			// txtProyeccionVentaDistribuidorSeptiembre
			// 
			appearance286.BackColorDisabled = System.Drawing.Color.White;
			appearance286.BackColorDisabled2 = System.Drawing.Color.White;
			appearance286.BorderColor = System.Drawing.Color.Black;
			appearance286.BorderColor3DBase = System.Drawing.Color.Black;
			appearance286.ForeColor = System.Drawing.Color.Black;
			appearance286.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorSeptiembre.Appearance = appearance286;
			this.txtProyeccionVentaDistribuidorSeptiembre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorSeptiembre.Enabled = false;
			this.txtProyeccionVentaDistribuidorSeptiembre.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorSeptiembre.Location = new System.Drawing.Point(762, 96);
			this.txtProyeccionVentaDistribuidorSeptiembre.Name = "txtProyeccionVentaDistribuidorSeptiembre";
			this.txtProyeccionVentaDistribuidorSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorSeptiembre.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorSeptiembre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorSeptiembre.TabIndex = 1149;
			// 
			// txtProyeccionVentaDistribuidorOctubre
			// 
			appearance287.BackColorDisabled = System.Drawing.Color.White;
			appearance287.BackColorDisabled2 = System.Drawing.Color.White;
			appearance287.BorderColor = System.Drawing.Color.Black;
			appearance287.BorderColor3DBase = System.Drawing.Color.Black;
			appearance287.ForeColor = System.Drawing.Color.Black;
			appearance287.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorOctubre.Appearance = appearance287;
			this.txtProyeccionVentaDistribuidorOctubre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorOctubre.Enabled = false;
			this.txtProyeccionVentaDistribuidorOctubre.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorOctubre.Location = new System.Drawing.Point(850, 96);
			this.txtProyeccionVentaDistribuidorOctubre.Name = "txtProyeccionVentaDistribuidorOctubre";
			this.txtProyeccionVentaDistribuidorOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorOctubre.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorOctubre.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorOctubre.TabIndex = 1148;
			// 
			// txtProyeccionOctubreVenta
			// 
			appearance288.ForeColor = System.Drawing.Color.Firebrick;
			appearance288.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionOctubreVenta.Appearance = appearance288;
			this.txtProyeccionOctubreVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionOctubreVenta.Enabled = false;
			this.txtProyeccionOctubreVenta.FormatString = "#,##0.00";
			this.txtProyeccionOctubreVenta.Location = new System.Drawing.Point(850, 56);
			this.txtProyeccionOctubreVenta.Name = "txtProyeccionOctubreVenta";
			this.txtProyeccionOctubreVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionOctubreVenta.PromptChar = ' ';
			this.txtProyeccionOctubreVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionOctubreVenta.TabIndex = 1147;
			// 
			// txtProyeccionSeptiembreVenta
			// 
			appearance289.ForeColor = System.Drawing.Color.Firebrick;
			appearance289.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionSeptiembreVenta.Appearance = appearance289;
			this.txtProyeccionSeptiembreVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionSeptiembreVenta.Enabled = false;
			this.txtProyeccionSeptiembreVenta.FormatString = "#,##0.00";
			this.txtProyeccionSeptiembreVenta.Location = new System.Drawing.Point(762, 56);
			this.txtProyeccionSeptiembreVenta.Name = "txtProyeccionSeptiembreVenta";
			this.txtProyeccionSeptiembreVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionSeptiembreVenta.PromptChar = ' ';
			this.txtProyeccionSeptiembreVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionSeptiembreVenta.TabIndex = 1146;
			// 
			// txtProyeccionVentaDistribuidorJulio
			// 
			appearance290.BackColorDisabled = System.Drawing.Color.White;
			appearance290.BackColorDisabled2 = System.Drawing.Color.White;
			appearance290.BorderColor = System.Drawing.Color.Black;
			appearance290.BorderColor3DBase = System.Drawing.Color.Black;
			appearance290.ForeColor = System.Drawing.Color.Black;
			appearance290.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorJulio.Appearance = appearance290;
			this.txtProyeccionVentaDistribuidorJulio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorJulio.Enabled = false;
			this.txtProyeccionVentaDistribuidorJulio.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorJulio.Location = new System.Drawing.Point(586, 96);
			this.txtProyeccionVentaDistribuidorJulio.Name = "txtProyeccionVentaDistribuidorJulio";
			this.txtProyeccionVentaDistribuidorJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorJulio.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorJulio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorJulio.TabIndex = 1145;
			// 
			// txtProyeccionVentaDistribuidorAgosto
			// 
			appearance291.BackColorDisabled = System.Drawing.Color.White;
			appearance291.BackColorDisabled2 = System.Drawing.Color.White;
			appearance291.BorderColor = System.Drawing.Color.Black;
			appearance291.BorderColor3DBase = System.Drawing.Color.Black;
			appearance291.ForeColor = System.Drawing.Color.Black;
			appearance291.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorAgosto.Appearance = appearance291;
			this.txtProyeccionVentaDistribuidorAgosto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorAgosto.Enabled = false;
			this.txtProyeccionVentaDistribuidorAgosto.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorAgosto.Location = new System.Drawing.Point(674, 96);
			this.txtProyeccionVentaDistribuidorAgosto.Name = "txtProyeccionVentaDistribuidorAgosto";
			this.txtProyeccionVentaDistribuidorAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorAgosto.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorAgosto.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorAgosto.TabIndex = 1144;
			// 
			// txtProyeccionAgostoVenta
			// 
			appearance292.ForeColor = System.Drawing.Color.Firebrick;
			appearance292.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionAgostoVenta.Appearance = appearance292;
			this.txtProyeccionAgostoVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionAgostoVenta.Enabled = false;
			this.txtProyeccionAgostoVenta.FormatString = "#,##0.00";
			this.txtProyeccionAgostoVenta.Location = new System.Drawing.Point(674, 56);
			this.txtProyeccionAgostoVenta.Name = "txtProyeccionAgostoVenta";
			this.txtProyeccionAgostoVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionAgostoVenta.PromptChar = ' ';
			this.txtProyeccionAgostoVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionAgostoVenta.TabIndex = 1143;
			// 
			// txtProyeccionJulioVenta
			// 
			appearance293.ForeColor = System.Drawing.Color.Firebrick;
			appearance293.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionJulioVenta.Appearance = appearance293;
			this.txtProyeccionJulioVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionJulioVenta.Enabled = false;
			this.txtProyeccionJulioVenta.FormatString = "#,##0.00";
			this.txtProyeccionJulioVenta.Location = new System.Drawing.Point(586, 56);
			this.txtProyeccionJulioVenta.Name = "txtProyeccionJulioVenta";
			this.txtProyeccionJulioVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionJulioVenta.PromptChar = ' ';
			this.txtProyeccionJulioVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionJulioVenta.TabIndex = 1142;
			// 
			// txtProyeccionVentaDistribuidorMayo
			// 
			appearance294.BackColorDisabled = System.Drawing.Color.White;
			appearance294.BackColorDisabled2 = System.Drawing.Color.White;
			appearance294.BorderColor = System.Drawing.Color.Black;
			appearance294.BorderColor3DBase = System.Drawing.Color.Black;
			appearance294.ForeColor = System.Drawing.Color.Black;
			appearance294.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorMayo.Appearance = appearance294;
			this.txtProyeccionVentaDistribuidorMayo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorMayo.Enabled = false;
			this.txtProyeccionVentaDistribuidorMayo.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorMayo.Location = new System.Drawing.Point(410, 96);
			this.txtProyeccionVentaDistribuidorMayo.Name = "txtProyeccionVentaDistribuidorMayo";
			this.txtProyeccionVentaDistribuidorMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorMayo.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorMayo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorMayo.TabIndex = 1141;
			// 
			// txtProyeccionVentaDistribuidorJunio
			// 
			appearance295.BackColorDisabled = System.Drawing.Color.White;
			appearance295.BackColorDisabled2 = System.Drawing.Color.White;
			appearance295.BorderColor = System.Drawing.Color.Black;
			appearance295.BorderColor3DBase = System.Drawing.Color.Black;
			appearance295.ForeColor = System.Drawing.Color.Black;
			appearance295.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorJunio.Appearance = appearance295;
			this.txtProyeccionVentaDistribuidorJunio.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorJunio.Enabled = false;
			this.txtProyeccionVentaDistribuidorJunio.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorJunio.Location = new System.Drawing.Point(498, 96);
			this.txtProyeccionVentaDistribuidorJunio.Name = "txtProyeccionVentaDistribuidorJunio";
			this.txtProyeccionVentaDistribuidorJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorJunio.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorJunio.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorJunio.TabIndex = 1140;
			// 
			// txtProyeccionJunioVenta
			// 
			appearance296.ForeColor = System.Drawing.Color.Firebrick;
			appearance296.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionJunioVenta.Appearance = appearance296;
			this.txtProyeccionJunioVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionJunioVenta.Enabled = false;
			this.txtProyeccionJunioVenta.FormatString = "#,##0.00";
			this.txtProyeccionJunioVenta.Location = new System.Drawing.Point(498, 56);
			this.txtProyeccionJunioVenta.Name = "txtProyeccionJunioVenta";
			this.txtProyeccionJunioVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionJunioVenta.PromptChar = ' ';
			this.txtProyeccionJunioVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionJunioVenta.TabIndex = 1139;
			// 
			// txtProyeccionMayoVenta
			// 
			appearance297.ForeColor = System.Drawing.Color.Firebrick;
			appearance297.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMayoVenta.Appearance = appearance297;
			this.txtProyeccionMayoVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMayoVenta.Enabled = false;
			this.txtProyeccionMayoVenta.FormatString = "#,##0.00";
			this.txtProyeccionMayoVenta.Location = new System.Drawing.Point(410, 56);
			this.txtProyeccionMayoVenta.Name = "txtProyeccionMayoVenta";
			this.txtProyeccionMayoVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMayoVenta.PromptChar = ' ';
			this.txtProyeccionMayoVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionMayoVenta.TabIndex = 1138;
			// 
			// txtProyeccionVentaDistribuidorMarzo
			// 
			appearance298.BackColorDisabled = System.Drawing.Color.White;
			appearance298.BackColorDisabled2 = System.Drawing.Color.White;
			appearance298.BorderColor = System.Drawing.Color.Black;
			appearance298.BorderColor3DBase = System.Drawing.Color.Black;
			appearance298.ForeColor = System.Drawing.Color.Black;
			appearance298.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorMarzo.Appearance = appearance298;
			this.txtProyeccionVentaDistribuidorMarzo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorMarzo.Enabled = false;
			this.txtProyeccionVentaDistribuidorMarzo.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorMarzo.Location = new System.Drawing.Point(234, 96);
			this.txtProyeccionVentaDistribuidorMarzo.Name = "txtProyeccionVentaDistribuidorMarzo";
			this.txtProyeccionVentaDistribuidorMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorMarzo.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorMarzo.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorMarzo.TabIndex = 1137;
			// 
			// txtProyeccionVentaDistribuidorAbril
			// 
			appearance299.BackColorDisabled = System.Drawing.Color.White;
			appearance299.BackColorDisabled2 = System.Drawing.Color.White;
			appearance299.BorderColor = System.Drawing.Color.Black;
			appearance299.BorderColor3DBase = System.Drawing.Color.Black;
			appearance299.ForeColor = System.Drawing.Color.Black;
			appearance299.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorAbril.Appearance = appearance299;
			this.txtProyeccionVentaDistribuidorAbril.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorAbril.Enabled = false;
			this.txtProyeccionVentaDistribuidorAbril.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorAbril.Location = new System.Drawing.Point(322, 96);
			this.txtProyeccionVentaDistribuidorAbril.Name = "txtProyeccionVentaDistribuidorAbril";
			this.txtProyeccionVentaDistribuidorAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorAbril.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorAbril.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorAbril.TabIndex = 1136;
			// 
			// txtProyeccionAbrilVenta
			// 
			appearance300.ForeColor = System.Drawing.Color.Firebrick;
			appearance300.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionAbrilVenta.Appearance = appearance300;
			this.txtProyeccionAbrilVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionAbrilVenta.Enabled = false;
			this.txtProyeccionAbrilVenta.FormatString = "#,##0.00";
			this.txtProyeccionAbrilVenta.Location = new System.Drawing.Point(322, 56);
			this.txtProyeccionAbrilVenta.Name = "txtProyeccionAbrilVenta";
			this.txtProyeccionAbrilVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionAbrilVenta.PromptChar = ' ';
			this.txtProyeccionAbrilVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionAbrilVenta.TabIndex = 1135;
			// 
			// txtProyeccionMarzoVenta
			// 
			appearance301.ForeColor = System.Drawing.Color.Firebrick;
			appearance301.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionMarzoVenta.Appearance = appearance301;
			this.txtProyeccionMarzoVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionMarzoVenta.Enabled = false;
			this.txtProyeccionMarzoVenta.FormatString = "#,##0.00";
			this.txtProyeccionMarzoVenta.Location = new System.Drawing.Point(234, 56);
			this.txtProyeccionMarzoVenta.Name = "txtProyeccionMarzoVenta";
			this.txtProyeccionMarzoVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionMarzoVenta.PromptChar = ' ';
			this.txtProyeccionMarzoVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionMarzoVenta.TabIndex = 1134;
			// 
			// txtProyeccionVentaDistribuidorEnero
			// 
			appearance302.BackColorDisabled = System.Drawing.Color.White;
			appearance302.BackColorDisabled2 = System.Drawing.Color.White;
			appearance302.BorderColor = System.Drawing.Color.Black;
			appearance302.BorderColor3DBase = System.Drawing.Color.Black;
			appearance302.ForeColor = System.Drawing.Color.Black;
			appearance302.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorEnero.Appearance = appearance302;
			this.txtProyeccionVentaDistribuidorEnero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorEnero.Enabled = false;
			this.txtProyeccionVentaDistribuidorEnero.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorEnero.Location = new System.Drawing.Point(58, 96);
			this.txtProyeccionVentaDistribuidorEnero.Name = "txtProyeccionVentaDistribuidorEnero";
			this.txtProyeccionVentaDistribuidorEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorEnero.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorEnero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorEnero.TabIndex = 1133;
			// 
			// txtProyeccionVentaDistribuidorFebrero
			// 
			appearance303.BackColorDisabled = System.Drawing.Color.White;
			appearance303.BackColorDisabled2 = System.Drawing.Color.White;
			appearance303.BorderColor = System.Drawing.Color.Black;
			appearance303.BorderColor3DBase = System.Drawing.Color.Black;
			appearance303.ForeColor = System.Drawing.Color.Black;
			appearance303.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyeccionVentaDistribuidorFebrero.Appearance = appearance303;
			this.txtProyeccionVentaDistribuidorFebrero.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProyeccionVentaDistribuidorFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionVentaDistribuidorFebrero.Enabled = false;
			this.txtProyeccionVentaDistribuidorFebrero.FormatString = "#,##0.00";
			this.txtProyeccionVentaDistribuidorFebrero.Location = new System.Drawing.Point(146, 96);
			this.txtProyeccionVentaDistribuidorFebrero.Name = "txtProyeccionVentaDistribuidorFebrero";
			this.txtProyeccionVentaDistribuidorFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionVentaDistribuidorFebrero.PromptChar = ' ';
			this.txtProyeccionVentaDistribuidorFebrero.Size = new System.Drawing.Size(80, 19);
			this.txtProyeccionVentaDistribuidorFebrero.TabIndex = 1132;
			// 
			// txtProyeccionFebreroVenta
			// 
			appearance304.ForeColor = System.Drawing.Color.Firebrick;
			appearance304.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionFebreroVenta.Appearance = appearance304;
			this.txtProyeccionFebreroVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionFebreroVenta.Enabled = false;
			this.txtProyeccionFebreroVenta.FormatString = "#,##0.00";
			this.txtProyeccionFebreroVenta.Location = new System.Drawing.Point(146, 56);
			this.txtProyeccionFebreroVenta.Name = "txtProyeccionFebreroVenta";
			this.txtProyeccionFebreroVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionFebreroVenta.PromptChar = ' ';
			this.txtProyeccionFebreroVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionFebreroVenta.TabIndex = 1131;
			// 
			// txtProyeccionEneroVenta
			// 
			appearance305.ForeColor = System.Drawing.Color.Firebrick;
			appearance305.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProyeccionEneroVenta.Appearance = appearance305;
			this.txtProyeccionEneroVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyeccionEneroVenta.Enabled = false;
			this.txtProyeccionEneroVenta.FormatString = "#,##0.00";
			this.txtProyeccionEneroVenta.Location = new System.Drawing.Point(58, 56);
			this.txtProyeccionEneroVenta.Name = "txtProyeccionEneroVenta";
			this.txtProyeccionEneroVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyeccionEneroVenta.PromptChar = ' ';
			this.txtProyeccionEneroVenta.Size = new System.Drawing.Size(80, 21);
			this.txtProyeccionEneroVenta.TabIndex = 1130;
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
			// txtAño
			// 
			this.txtAño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAño.Location = new System.Drawing.Point(128, 10);
			this.txtAño.MinValue = 2017;
			this.txtAño.Name = "txtAño";
			this.txtAño.PromptChar = ' ';
			this.txtAño.Size = new System.Drawing.Size(88, 21);
			this.txtAño.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAño.TabIndex = 1456;
			this.txtAño.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAño_KeyDown);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.Location = new System.Drawing.Point(744, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(40, 24);
			this.btnVer.TabIndex = 1457;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.Location = new System.Drawing.Point(240, 12);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(31, 16);
			this.label27.TabIndex = 1458;
			this.label27.Text = "Local";
			this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource2;
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 180;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(280, 10);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 21);
			this.cmbBodega.TabIndex = 1459;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			// 
			// ultraDataSource2
			// 
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// label41
			// 
			this.label41.AutoSize = true;
			this.label41.Location = new System.Drawing.Point(56, 12);
			this.label41.Name = "label41";
			this.label41.Size = new System.Drawing.Size(43, 16);
			this.label41.TabIndex = 1460;
			this.label41.Text = "Periodo";
			// 
			// optContrato
			// 
			appearance306.ForeColorDisabled = System.Drawing.Color.Black;
			this.optContrato.Appearance = appearance306;
			this.optContrato.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optContrato.CheckedIndex = 0;
			this.optContrato.ItemAppearance = appearance307;
			this.optContrato.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = false;
			valueListItem1.DisplayText = "Propios";
			valueListItem2.DataValue = true;
			valueListItem2.DisplayText = "Generales";
			this.optContrato.Items.Add(valueListItem1);
			this.optContrato.Items.Add(valueListItem2);
			this.optContrato.ItemSpacingVertical = 10;
			this.optContrato.Location = new System.Drawing.Point(520, 8);
			this.optContrato.Name = "optContrato";
			this.optContrato.Size = new System.Drawing.Size(192, 24);
			this.optContrato.TabIndex = 1461;
			this.optContrato.Text = "Propios";
			this.optContrato.ValueChanged += new System.EventHandler(this.optContrato_ValueChanged);
			// 
			// frmProyeccionPresupuestoAlmacen
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1222, 734);
			this.Controls.Add(this.optContrato);
			this.Controls.Add(this.label41);
			this.Controls.Add(this.label27);
			this.Controls.Add(this.label40);
			this.Controls.Add(this.label39);
			this.Controls.Add(this.label38);
			this.Controls.Add(this.label37);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.label35);
			this.Controls.Add(this.label33);
			this.Controls.Add(this.label34);
			this.Controls.Add(this.label32);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.label30);
			this.Controls.Add(this.label29);
			this.Controls.Add(this.label28);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtAño);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.txtTotalGeneral);
			this.Controls.Add(this.txtTotalGeneralNoviembre);
			this.Controls.Add(this.txtTotalGeneralDiciembre);
			this.Controls.Add(this.txtTotalGeneralSeptiembre);
			this.Controls.Add(this.txtTotalGeneralOctubre);
			this.Controls.Add(this.txtTotalGeneralJulio);
			this.Controls.Add(this.txtTotalGeneralAgosto);
			this.Controls.Add(this.txtTotalGeneralMayo);
			this.Controls.Add(this.txtTotalGeneralJunio);
			this.Controls.Add(this.txtTotalGeneralMarzo);
			this.Controls.Add(this.txtTotalGeneralAbril);
			this.Controls.Add(this.txtTotalGeneralFebrero);
			this.Controls.Add(this.txtTotalGeneralEnero);
			this.Controls.Add(this.txtFDGTotal);
			this.Controls.Add(this.txtFDGNoviembre);
			this.Controls.Add(this.txtFDGDiciembre);
			this.Controls.Add(this.txtFDGSeptiembre);
			this.Controls.Add(this.txtFDGOctubre);
			this.Controls.Add(this.txtFDGJulio);
			this.Controls.Add(this.txtFDGAgosto);
			this.Controls.Add(this.txtFDGMayo);
			this.Controls.Add(this.txtFDGJunio);
			this.Controls.Add(this.txtFDGMarzo);
			this.Controls.Add(this.txtFDGAbril);
			this.Controls.Add(this.txtFDGFebrero);
			this.Controls.Add(this.txtFDGEnero);
			this.Controls.Add(this.txtPrestamosTotal);
			this.Controls.Add(this.txtPrestamosNoviembre);
			this.Controls.Add(this.txtPrestamosDiciembre);
			this.Controls.Add(this.txtPrestamosSeptiembre);
			this.Controls.Add(this.txtPrestamosOctubre);
			this.Controls.Add(this.txtPrestamosJulio);
			this.Controls.Add(this.txtPrestamosAgosto);
			this.Controls.Add(this.txtPrestamosMayo);
			this.Controls.Add(this.txtPrestamosJunio);
			this.Controls.Add(this.txtPrestamosMarzo);
			this.Controls.Add(this.txtPrestamosAbril);
			this.Controls.Add(this.txtPrestamosFebrero);
			this.Controls.Add(this.txtPrestamosEnero);
			this.Controls.Add(this.uGridGastos);
			this.Controls.Add(this.txtGNCTotal);
			this.Controls.Add(this.txtGNCNoviembre);
			this.Controls.Add(this.txtGNCDiciembre);
			this.Controls.Add(this.txtGNCSeptiembre);
			this.Controls.Add(this.txtGNCOctubre);
			this.Controls.Add(this.txtGNCJulio);
			this.Controls.Add(this.txtGNCAgosto);
			this.Controls.Add(this.txtGNCMayo);
			this.Controls.Add(this.txtGNCJunio);
			this.Controls.Add(this.txtGNCMarzo);
			this.Controls.Add(this.txtGNCAbril);
			this.Controls.Add(this.txtGNCFebrero);
			this.Controls.Add(this.txtGNCEnero);
			this.Controls.Add(this.txtMNCTotal);
			this.Controls.Add(this.txtMNCNoviembre);
			this.Controls.Add(this.txtMNCDiciembre);
			this.Controls.Add(this.txtMNCSeptiembre);
			this.Controls.Add(this.txtMNCOctubre);
			this.Controls.Add(this.txtMNCJulio);
			this.Controls.Add(this.txtMNCAgosto);
			this.Controls.Add(this.txtMNCMayo);
			this.Controls.Add(this.txtMNCJunio);
			this.Controls.Add(this.txtMNCMarzo);
			this.Controls.Add(this.txtMNCAbril);
			this.Controls.Add(this.txtMNCFebrero);
			this.Controls.Add(this.txtMNCEnero);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.txtSMAGastosRetirosDiciembre);
			this.Controls.Add(this.txtSMAGastosRetirosNoviembre);
			this.Controls.Add(this.txtSMAGastosRetirosOctubre);
			this.Controls.Add(this.txtSMAGastosRetirosSeptiembre);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.txtSMAGastosRetirosAgosto);
			this.Controls.Add(this.txtSMAGastosRetirosJulio);
			this.Controls.Add(this.txtSMAGastosRetirosJunio);
			this.Controls.Add(this.txtSMAGastosRetirosMayo);
			this.Controls.Add(this.txtSMAGastosRetirosAbril);
			this.Controls.Add(this.txtSMAGastosRetirosMarzo);
			this.Controls.Add(this.txtSMAGastosRetirosFebrero);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.txtTotalGastos);
			this.Controls.Add(this.txtTotalGastosNoviembre);
			this.Controls.Add(this.txtTotalGastosDiciembre);
			this.Controls.Add(this.txtTotalGastosSeptiembre);
			this.Controls.Add(this.txtTotalGastosOctubre);
			this.Controls.Add(this.txtTotalGastosJulio);
			this.Controls.Add(this.txtTotalGastosAgosto);
			this.Controls.Add(this.txtTotalGastosMayo);
			this.Controls.Add(this.txtTotalGastosJunio);
			this.Controls.Add(this.txtTotalGastosMarzo);
			this.Controls.Add(this.txtTotalGastosAbril);
			this.Controls.Add(this.txtTotalGastosFebrero);
			this.Controls.Add(this.txtTotalGastosEnero);
			this.Controls.Add(this.txtResultadoTotal);
			this.Controls.Add(this.txtResultadoNoviembre);
			this.Controls.Add(this.txtResultadoDiciembre);
			this.Controls.Add(this.txtResultadoSeptiembre);
			this.Controls.Add(this.txtResultadoOctubre);
			this.Controls.Add(this.txtResultadoJulio);
			this.Controls.Add(this.txtResultadoAgosto);
			this.Controls.Add(this.txtResultadoMayo);
			this.Controls.Add(this.txtResultadoJunio);
			this.Controls.Add(this.txtResultadoMarzo);
			this.Controls.Add(this.txtResultadoAbril);
			this.Controls.Add(this.txtResultadoFebrero);
			this.Controls.Add(this.txtResultadoEnero);
			this.Controls.Add(this.txtGastosRetiroTotal);
			this.Controls.Add(this.txtGastosRetiroNoviembre);
			this.Controls.Add(this.txtGastosRetiroDiciembre);
			this.Controls.Add(this.txtGastosRetiroSeptiembre);
			this.Controls.Add(this.txtGastosRetiroOctubre);
			this.Controls.Add(this.txtGastosRetiroJulio);
			this.Controls.Add(this.txtGastosRetiroAgosto);
			this.Controls.Add(this.txtGastosRetiroMayo);
			this.Controls.Add(this.txtGastosRetiroJunio);
			this.Controls.Add(this.txtGastosRetiroMarzo);
			this.Controls.Add(this.txtGastosRetiroAbril);
			this.Controls.Add(this.txtGastosRetiroFebrero);
			this.Controls.Add(this.txtGastosRetiroEnero);
			this.Controls.Add(this.txtTotalIngresos);
			this.Controls.Add(this.txtTotalIngresosNoviembre);
			this.Controls.Add(this.txtTotalIngresosDiciembre);
			this.Controls.Add(this.txtTotalIngresosSeptiembre);
			this.Controls.Add(this.txtTotalIngresosOctubre);
			this.Controls.Add(this.txtTotalIngresosJulio);
			this.Controls.Add(this.txtTotalIngresosAgosto);
			this.Controls.Add(this.txtTotalIngresosMayo);
			this.Controls.Add(this.txtTotalIngresosJunio);
			this.Controls.Add(this.txtTotalIngresosMarzo);
			this.Controls.Add(this.txtTotalIngresosAbril);
			this.Controls.Add(this.txtTotalIngresosFebrero);
			this.Controls.Add(this.txtTotalIngresosEnero);
			this.Controls.Add(this.txtCobranzaTotal);
			this.Controls.Add(this.txtCobranzaNoviembre);
			this.Controls.Add(this.txtCobranzaDiciembre);
			this.Controls.Add(this.txtCobranzaSeptiembre);
			this.Controls.Add(this.txtCobranzaOctubre);
			this.Controls.Add(this.txtCobranzaJulio);
			this.Controls.Add(this.txtCobranzaAgosto);
			this.Controls.Add(this.txtCobranzaMayo);
			this.Controls.Add(this.txtCobranzaJunio);
			this.Controls.Add(this.txtCobranzaMarzo);
			this.Controls.Add(this.txtCobranzaAbril);
			this.Controls.Add(this.txtCobranzaFebrero);
			this.Controls.Add(this.txtCobranzaEnero);
			this.Controls.Add(this.txtProyeccionMargenTotal);
			this.Controls.Add(this.txtProyeccionMargenTotalNoviembre);
			this.Controls.Add(this.txtProyeccionMargenTotalDiciembre);
			this.Controls.Add(this.txtProyeccionMargenTotalSeptiembre);
			this.Controls.Add(this.txtProyeccionMargenTotalOctubre);
			this.Controls.Add(this.txtProyeccionMargenTotalJulio);
			this.Controls.Add(this.txtProyeccionMargenTotalAgosto);
			this.Controls.Add(this.txtProyeccionMargenTotalMayo);
			this.Controls.Add(this.txtProyeccionMargenTotalJunio);
			this.Controls.Add(this.txtProyeccionMargenTotalMarzo);
			this.Controls.Add(this.txtProyeccionMargenTotalAbril);
			this.Controls.Add(this.txtProyeccionMargenTotalFebrero);
			this.Controls.Add(this.txtProyeccionMargenTotalEnero);
			this.Controls.Add(this.label25);
			this.Controls.Add(this.txtProyeccionVentaTotal);
			this.Controls.Add(this.txtProyeccionVentaTotalNoviembre);
			this.Controls.Add(this.txtProyeccionVentaTotalDiciembre);
			this.Controls.Add(this.txtProyeccionVentaTotalSeptiembre);
			this.Controls.Add(this.txtProyeccionVentaTotalOctubre);
			this.Controls.Add(this.txtProyeccionVentaTotalJulio);
			this.Controls.Add(this.txtProyeccionVentaTotalAgosto);
			this.Controls.Add(this.txtProyeccionVentaTotalMayo);
			this.Controls.Add(this.txtProyeccionVentaTotalJunio);
			this.Controls.Add(this.txtProyeccionVentaTotalMarzo);
			this.Controls.Add(this.txtProyeccionVentaTotalAbril);
			this.Controls.Add(this.txtProyeccionVentaTotalFebrero);
			this.Controls.Add(this.txtProyeccionVentaTotalEnero);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtProyeccionMargenCredito);
			this.Controls.Add(this.txtProyeccionMargenCreditoNoviembre);
			this.Controls.Add(this.txtProyeccionMargenCreditoDiciembre);
			this.Controls.Add(this.txtProyeccionMargenCreditoSeptiembre);
			this.Controls.Add(this.txtProyeccionMargenCreditoOctubre);
			this.Controls.Add(this.txtProyeccionMargenCreditoJulio);
			this.Controls.Add(this.txtProyeccionMargenCreditoAgosto);
			this.Controls.Add(this.txtProyeccionMargenCreditoMayo);
			this.Controls.Add(this.txtProyeccionMargenCreditoJunio);
			this.Controls.Add(this.txtProyeccionMargenCreditoMarzo);
			this.Controls.Add(this.txtProyeccionMargenCreditoAbril);
			this.Controls.Add(this.txtProyeccionMargenCreditoEnero);
			this.Controls.Add(this.txtProyeccionMargenCreditoFebrero);
			this.Controls.Add(this.txtProyeccionMargenTarjeta);
			this.Controls.Add(this.txtProyeccionMargenTarjetaNoviembre);
			this.Controls.Add(this.txtProyeccionMargenTarjetaDiciembre);
			this.Controls.Add(this.txtProyeccionMargenTarjetaSeptiembre);
			this.Controls.Add(this.txtProyeccionMargenTarjetaOctubre);
			this.Controls.Add(this.txtProyeccionMargenTarjetaJulio);
			this.Controls.Add(this.txtProyeccionMargenTarjetaAgosto);
			this.Controls.Add(this.txtProyeccionMargenTarjetaMayo);
			this.Controls.Add(this.txtProyeccionMargenTarjetaJunio);
			this.Controls.Add(this.txtProyeccionMargenTarjetaMarzo);
			this.Controls.Add(this.txtProyeccionMargenTarjetaAbril);
			this.Controls.Add(this.txtProyeccionMargenTarjetaEnero);
			this.Controls.Add(this.txtProyeccionMargenTarjetaFebrero);
			this.Controls.Add(this.txtProyeccionMargenEfectivo);
			this.Controls.Add(this.txtProyeccionMargenEfectivoNoviembre);
			this.Controls.Add(this.txtProyeccionMargenEfectivoDiciembre);
			this.Controls.Add(this.txtProyeccionMargenEfectivoSeptiembre);
			this.Controls.Add(this.txtProyeccionMargenEfectivoOctubre);
			this.Controls.Add(this.txtProyeccionMargenEfectivoJulio);
			this.Controls.Add(this.txtProyeccionMargenEfectivoAgosto);
			this.Controls.Add(this.txtProyeccionMargenEfectivoMayo);
			this.Controls.Add(this.txtProyeccionMargenEfectivoJunio);
			this.Controls.Add(this.txtProyeccionMargenEfectivoMarzo);
			this.Controls.Add(this.txtProyeccionMargenEfectivoAbril);
			this.Controls.Add(this.txtProyeccionMargenEfectivoEnero);
			this.Controls.Add(this.txtProyeccionMargenEfectivoFebrero);
			this.Controls.Add(this.txtProyeccionMargenDistribuidor);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorNoviembre);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorDiciembre);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorSeptiembre);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorOctubre);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorJulio);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorAgosto);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorMayo);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorJunio);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorMarzo);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorAbril);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorEnero);
			this.Controls.Add(this.txtProyeccionMargenDistribuidorFebrero);
			this.Controls.Add(this.txtProyeccionVentaCredito);
			this.Controls.Add(this.txtProyeccionVentaCreditoNoviembre);
			this.Controls.Add(this.txtProyeccionVentaCreditoDiciembre);
			this.Controls.Add(this.txtProyeccionVentaCreditoSeptiembre);
			this.Controls.Add(this.txtProyeccionVentaCreditoOctubre);
			this.Controls.Add(this.txtProyeccionVentaCreditoJulio);
			this.Controls.Add(this.txtProyeccionVentaCreditoAgosto);
			this.Controls.Add(this.txtProyeccionVentaCreditoMayo);
			this.Controls.Add(this.txtProyeccionVentaCreditoJunio);
			this.Controls.Add(this.txtProyeccionVentaCreditoMarzo);
			this.Controls.Add(this.txtProyeccionVentaCreditoAbril);
			this.Controls.Add(this.txtProyeccionVentaCreditoEnero);
			this.Controls.Add(this.txtProyeccionVentaCreditoFebrero);
			this.Controls.Add(this.txtProyeccionVentaTarjeta);
			this.Controls.Add(this.txtProyeccionVentaTarjetaNoviembre);
			this.Controls.Add(this.txtProyeccionVentaTarjetaDiciembre);
			this.Controls.Add(this.txtProyeccionVentaTarjetaSeptiembre);
			this.Controls.Add(this.txtProyeccionVentaTarjetaOctubre);
			this.Controls.Add(this.txtProyeccionVentaTarjetaJulio);
			this.Controls.Add(this.txtProyeccionVentaTarjetaAgosto);
			this.Controls.Add(this.txtProyeccionVentaTarjetaMayo);
			this.Controls.Add(this.txtProyeccionVentaTarjetaJunio);
			this.Controls.Add(this.txtProyeccionVentaTarjetaMarzo);
			this.Controls.Add(this.txtProyeccionVentaTarjetaAbril);
			this.Controls.Add(this.txtProyeccionVentaTarjetaEnero);
			this.Controls.Add(this.txtProyeccionVentaTarjetaFebrero);
			this.Controls.Add(this.txtProyeccionVentaEfectivo);
			this.Controls.Add(this.txtProyeccionVentaEfectivoNoviembre);
			this.Controls.Add(this.txtProyeccionVentaEfectivoDiciembre);
			this.Controls.Add(this.txtProyeccionVentaEfectivoSeptiembre);
			this.Controls.Add(this.txtProyeccionVentaEfectivoOctubre);
			this.Controls.Add(this.txtProyeccionVentaEfectivoJulio);
			this.Controls.Add(this.txtProyeccionVentaEfectivoAgosto);
			this.Controls.Add(this.txtProyeccionVentaEfectivoMayo);
			this.Controls.Add(this.txtProyeccionVentaEfectivoJunio);
			this.Controls.Add(this.txtProyeccionVentaEfectivoMarzo);
			this.Controls.Add(this.txtProyeccionVentaEfectivoAbril);
			this.Controls.Add(this.txtProyeccionVentaEfectivoEnero);
			this.Controls.Add(this.txtProyeccionVentaEfectivoFebrero);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtProyeccionVentaDistribuidor);
			this.Controls.Add(this.txtIngresosTotal);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorNoviembre);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorDiciembre);
			this.Controls.Add(this.txtProyeccionDiciembreVenta);
			this.Controls.Add(this.txtProyeccionNoviembreVenta);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorSeptiembre);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorOctubre);
			this.Controls.Add(this.txtProyeccionOctubreVenta);
			this.Controls.Add(this.txtProyeccionSeptiembreVenta);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorJulio);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorAgosto);
			this.Controls.Add(this.txtProyeccionAgostoVenta);
			this.Controls.Add(this.txtProyeccionJulioVenta);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorMayo);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorJunio);
			this.Controls.Add(this.txtProyeccionJunioVenta);
			this.Controls.Add(this.txtProyeccionMayoVenta);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorMarzo);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorAbril);
			this.Controls.Add(this.txtProyeccionAbrilVenta);
			this.Controls.Add(this.txtProyeccionMarzoVenta);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorEnero);
			this.Controls.Add(this.txtProyeccionVentaDistribuidorFebrero);
			this.Controls.Add(this.txtProyeccionFebreroVenta);
			this.Controls.Add(this.txtProyeccionEneroVenta);
			this.KeyPreview = true;
			this.Name = "frmProyeccionPresupuestoAlmacen";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Presupuesto Por Almacen";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmProyeccionPresupuestoAlmacen_KeyDown);
			this.Load += new System.EventHandler(this.frmProyeccionPresupuestoAlmacen_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneral)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneralEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFDGEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrestamosEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGNCEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMNCEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSMAGastosRetirosFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastosEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultadoEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosRetiroEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresosEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobranzaEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTotalEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTotalEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenCreditoFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenTarjetaFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenEfectivoFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMargenDistribuidorFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaCreditoFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaTarjetaFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaEfectivoFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionDiciembreVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionNoviembreVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionOctubreVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionSeptiembreVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionAgostoVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionJulioVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionJunioVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMayoVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionAbrilVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionMarzoVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionVentaDistribuidorFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionFebreroVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyeccionEneroVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optContrato)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		

		private void Consulta ()
		{
			int IBodega = 0;
			if (this.cmbBodega.ActiveRow != null) IBodega = (int)this.cmbBodega.Value;
			dtFechaInicial = DateTime.Parse("01/01/" + iPeriodo.ToString());
			dtFechaFinal = DateTime.Parse("31/12/" + iPeriodo.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualGastosAlmacenes '{0}', '{1}', {2},{3}", 
				dtFechaInicial.ToString("yyyyMMdd"), dtFechaFinal.ToString("yyyyMMdd"), iPeriodo, IBodega);
			this.uGridGastos.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			DatosEnero();
			DatosFebrero();
			DatosMarzo();
			DatosAbril();
			DatosMayo();
			DatosJunio();
			DatosJulio();
			DatosAgosto();
			DatosSeptiembre();
			DatosOctubre();
			DatosNoviembre();
			DatosDiciembre();
		}

		private void ConsultaGeneral()
		{
			int IBodega = 0;
			if (this.cmbBodega.ActiveRow != null) IBodega = (int)this.cmbBodega.Value;
			dtFechaInicial = DateTime.Parse("01/01/" + iPeriodo.ToString());
			dtFechaFinal = DateTime.Parse("31/12/" + iPeriodo.ToString());

			DataSet oDSConsulta = new DataSet();
			
			cdsSeteoF.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = cdsSeteoF.Schema.Connections[0];
			miConeccion = cdsSeteoF.Schema.Connections[0];

			string sSQL = string.Format("Exec ProyeccionMensualVentasLocales '{0}', '{1}',{2}", 
				dtFechaInicial.ToString("yyyyMMdd"), dtFechaFinal.ToString("yyyyMMdd"), IBodega);
			SqlDataAdapter oDAConsulta = new SqlDataAdapter(sSQL, (SqlConnection) miConeccion.DbConnection);
			oDAConsulta.Fill(oDSConsulta);

			#region Carga Datos
			foreach(DataRow oDRConsulta in oDSConsulta.Tables[0].Rows)
			{
				#region Proyeccion de Ventas
				if ((int)oDRConsulta["idTipo"] == 1)
				{
					if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtProyeccionEneroVenta.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
					if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtProyeccionFebreroVenta.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
					if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtProyeccionMarzoVenta.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
					if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtProyeccionAbrilVenta.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
					if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtProyeccionMayoVenta.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
					if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtProyeccionJunioVenta.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
					if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtProyeccionJulioVenta.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
					if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtProyeccionAgostoVenta.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
					if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtProyeccionSeptiembreVenta.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
					if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtProyeccionOctubreVenta.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
					if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtProyeccionNoviembreVenta.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
					if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtProyeccionDiciembreVenta.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

					this.txtIngresosTotal.Value = 
						Convert.ToDecimal(this.txtProyeccionEneroVenta.Value) + 
						Convert.ToDecimal(this.txtProyeccionFebreroVenta.Value) + 
						Convert.ToDecimal(this.txtProyeccionMarzoVenta.Value) + 
						Convert.ToDecimal(this.txtProyeccionAbrilVenta.Value) + 
						Convert.ToDecimal(this.txtProyeccionMayoVenta.Value) + 
						Convert.ToDecimal(this.txtProyeccionJunioVenta.Value) + 
						Convert.ToDecimal(this.txtProyeccionJulioVenta.Value) + 
						Convert.ToDecimal(this.txtProyeccionAgostoVenta.Value) + 
						Convert.ToDecimal(this.txtProyeccionSeptiembreVenta.Value) + 
						Convert.ToDecimal(this.txtProyeccionOctubreVenta.Value) + 
						Convert.ToDecimal(this.txtProyeccionNoviembreVenta.Value) + 
						Convert.ToDecimal(this.txtProyeccionDiciembreVenta.Value);
				}
				#endregion Proyeccion de Ventas

				#region Ventas
				if ((int)oDRConsulta["idTipo"] == 2)
				{
					#region Ventas Distribuidor
					if (oDRConsulta["FormaDePago"].ToString() == "DISTRIBUIDOR")
					{
						if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
						if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
						if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
						if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
						if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
						if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
						if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
						if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
						if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
						if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
						if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
						if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtProyeccionVentaDistribuidorDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

						this.txtProyeccionVentaDistribuidor.Value = 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorEnero.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorFebrero.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorMarzo.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorAbril.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorMayo.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorJunio.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorJulio.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorAgosto.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorSeptiembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorOctubre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorNoviembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaDistribuidorDiciembre.Value);
					}
					#endregion Ventas Distribuidor

					#region Ventas Efectivo
					if (oDRConsulta["FormaDePago"].ToString() == "EFECTIVO")
					{
						if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
						if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
						if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
						if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
						if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
						if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
						if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
						if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
						if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
						if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
						if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
						if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtProyeccionVentaEfectivoDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

						this.txtProyeccionVentaEfectivo.Value = 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoEnero.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoFebrero.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoMarzo.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoAbril.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoMayo.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoJunio.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoJulio.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoAgosto.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoSeptiembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoOctubre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoNoviembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaEfectivoDiciembre.Value);
					}
					#endregion Ventas Efectivo
					
					#region Ventas Tarjeta
					if (oDRConsulta["FormaDePago"].ToString() == "TARJETA")
					{
						if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
						if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
						if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
						if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
						if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
						if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
						if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
						if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
						if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
						if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
						if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
						if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtProyeccionVentaTarjetaDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

						this.txtProyeccionVentaTarjeta.Value = 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaEnero.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaFebrero.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaMarzo.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaAbril.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaMayo.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaJunio.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaJulio.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaAgosto.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaSeptiembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaOctubre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaNoviembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaTarjetaDiciembre.Value);
					}
					#endregion Ventas Tarjeta

					#region Ventas Credito
					if (oDRConsulta["FormaDePago"].ToString() == "CREDITO")
					{
						if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtProyeccionVentaCreditoEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
						if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtProyeccionVentaCreditoFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
						if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtProyeccionVentaCreditoMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
						if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtProyeccionVentaCreditoAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
						if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtProyeccionVentaCreditoMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
						if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtProyeccionVentaCreditoJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
						if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtProyeccionVentaCreditoJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
						if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtProyeccionVentaCreditoAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
						if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtProyeccionVentaCreditoSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
						if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtProyeccionVentaCreditoOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
						if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtProyeccionVentaCreditoNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
						if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtProyeccionVentaCreditoDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

						this.txtProyeccionVentaCredito.Value = 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoEnero.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoFebrero.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoMarzo.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoAbril.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoMayo.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoJunio.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoJulio.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoAgosto.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoSeptiembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoOctubre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoNoviembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionVentaCreditoDiciembre.Value);
					}
					#endregion Ventas Credito
				}
				#endregion Ventas
									
				#region Margen
				if ((int)oDRConsulta["idTipo"] == 3)
				{
					#region Margen Distribuidor
					if (oDRConsulta["FormaDePago"].ToString() == "DISTRIBUIDOR")
					{
						if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
						if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
						if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
						if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
						if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
						if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
						if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
						if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
						if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
						if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
						if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
						if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtProyeccionMargenDistribuidorDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

						this.txtProyeccionMargenDistribuidor.Value = 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorEnero.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorFebrero.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorMarzo.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorAbril.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorMayo.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorJunio.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorJulio.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorAgosto.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorSeptiembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorOctubre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorNoviembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenDistribuidorDiciembre.Value);
					}
					#endregion Margen Distribuidor

					#region Margen Efectivo
					if (oDRConsulta["FormaDePago"].ToString() == "EFECTIVO")
					{
						if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
						if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
						if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
						if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
						if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
						if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
						if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
						if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
						if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
						if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
						if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
						if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtProyeccionMargenEfectivoDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

						this.txtProyeccionMargenEfectivo.Value = 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoEnero.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoFebrero.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoMarzo.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoAbril.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoMayo.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoJunio.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoJulio.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoAgosto.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoSeptiembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoOctubre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoNoviembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenEfectivoDiciembre.Value);
					}
					#endregion Margen Efectivo

					#region Margen Tarjeta
					if (oDRConsulta["FormaDePago"].ToString() == "TARJETA")
					{
						if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
						if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
						if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
						if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
						if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
						if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
						if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
						if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
						if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
						if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
						if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
						if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtProyeccionMargenTarjetaDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

						this.txtProyeccionMargenTarjeta.Value = 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaEnero.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaFebrero.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaMarzo.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaAbril.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaMayo.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaJunio.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaJulio.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaAgosto.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaSeptiembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaOctubre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaNoviembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenTarjetaDiciembre.Value);
					}
					#endregion Margen Tarjeta

					#region Margen Credito
					if (oDRConsulta["FormaDePago"].ToString() == "CREDITO")
					{
						if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtProyeccionMargenCreditoEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
						if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtProyeccionMargenCreditoFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
						if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtProyeccionMargenCreditoMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
						if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtProyeccionMargenCreditoAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
						if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtProyeccionMargenCreditoMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
						if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtProyeccionMargenCreditoJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
						if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtProyeccionMargenCreditoJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
						if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtProyeccionMargenCreditoAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
						if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtProyeccionMargenCreditoSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
						if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtProyeccionMargenCreditoOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
						if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtProyeccionMargenCreditoNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
						if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtProyeccionMargenCreditoDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

						this.txtProyeccionMargenCredito.Value = 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoEnero.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoFebrero.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoMarzo.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoAbril.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoMayo.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoJunio.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoJulio.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoAgosto.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoSeptiembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoOctubre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoNoviembre.Value) + 
							Convert.ToDecimal(this.txtProyeccionMargenCreditoDiciembre.Value);
					}
					#endregion Margen Credito
				}
				#endregion Margen

				#region Cobranza
				if ((int)oDRConsulta["idTipo"] == 4)
				{
					if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtCobranzaEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
					if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtCobranzaFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
					if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtCobranzaMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
					if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtCobranzaAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
					if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtCobranzaMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
					if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtCobranzaJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
					if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtCobranzaJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
					if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtCobranzaAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
					if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtCobranzaSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
					if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtCobranzaOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
					if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtCobranzaNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
					if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtCobranzaDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

					this.txtCobranzaTotal.Value = 
						Convert.ToDecimal(this.txtCobranzaEnero.Value) + 
						Convert.ToDecimal(this.txtCobranzaFebrero.Value) + 
						Convert.ToDecimal(this.txtCobranzaMarzo.Value) + 
						Convert.ToDecimal(this.txtCobranzaAbril.Value) + 
						Convert.ToDecimal(this.txtCobranzaMayo.Value) + 
						Convert.ToDecimal(this.txtCobranzaJunio.Value) + 
						Convert.ToDecimal(this.txtCobranzaJulio.Value) + 
						Convert.ToDecimal(this.txtCobranzaAgosto.Value) + 
						Convert.ToDecimal(this.txtCobranzaSeptiembre.Value) + 
						Convert.ToDecimal(this.txtCobranzaOctubre.Value) + 
						Convert.ToDecimal(this.txtCobranzaNoviembre.Value) + 
						Convert.ToDecimal(this.txtCobranzaDiciembre.Value);
				}
				#endregion Cobranza

				#region Retiros
				if ((int)oDRConsulta["idTipo"] == 5)
				{
					if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtGastosRetiroEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]) * - 1;
					if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtGastosRetiroFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]) * - 1;
					if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtGastosRetiroMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]) * - 1;
					if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtGastosRetiroAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]) * - 1;
					if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtGastosRetiroMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]) * - 1;
					if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtGastosRetiroJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]) * - 1;
					if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtGastosRetiroJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]) * - 1;
					if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtGastosRetiroAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]) * - 1;
					if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtGastosRetiroSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]) * - 1;
					if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtGastosRetiroOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]) * - 1;
					if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtGastosRetiroNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]) * - 1;
					if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtGastosRetiroDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]) * - 1;

					this.txtGastosRetiroTotal.Value = 
						Convert.ToDecimal(this.txtGastosRetiroEnero.Value) + 
						Convert.ToDecimal(this.txtGastosRetiroFebrero.Value) + 
						Convert.ToDecimal(this.txtGastosRetiroMarzo.Value) + 
						Convert.ToDecimal(this.txtGastosRetiroAbril.Value) + 
						Convert.ToDecimal(this.txtGastosRetiroMayo.Value) + 
						Convert.ToDecimal(this.txtGastosRetiroJunio.Value) + 
						Convert.ToDecimal(this.txtGastosRetiroJulio.Value) + 
						Convert.ToDecimal(this.txtGastosRetiroAgosto.Value) + 
						Convert.ToDecimal(this.txtGastosRetiroSeptiembre.Value) + 
						Convert.ToDecimal(this.txtGastosRetiroOctubre.Value) + 
						Convert.ToDecimal(this.txtGastosRetiroNoviembre.Value) + 
						Convert.ToDecimal(this.txtGastosRetiroDiciembre.Value);
				}
				#endregion Retiros

				#region Gastos Notas De Credito
				if ((int)oDRConsulta["idTipo"] == 6)
				{
					if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtGNCEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
					if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtGNCFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
					if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtGNCMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
					if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtGNCAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
					if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtGNCMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
					if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtGNCJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
					if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtGNCJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
					if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtGNCAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
					if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtGNCSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
					if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtGNCOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
					if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtGNCNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
					if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtGNCDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

					this.txtGNCTotal.Value = 
						Convert.ToDecimal(this.txtGNCEnero.Value) + 
						Convert.ToDecimal(this.txtGNCFebrero.Value) + 
						Convert.ToDecimal(this.txtGNCMarzo.Value) + 
						Convert.ToDecimal(this.txtGNCAbril.Value) + 
						Convert.ToDecimal(this.txtGNCMayo.Value) + 
						Convert.ToDecimal(this.txtGNCJunio.Value) + 
						Convert.ToDecimal(this.txtGNCJulio.Value) + 
						Convert.ToDecimal(this.txtGNCAgosto.Value) + 
						Convert.ToDecimal(this.txtGNCSeptiembre.Value) + 
						Convert.ToDecimal(this.txtGNCOctubre.Value) + 
						Convert.ToDecimal(this.txtGNCNoviembre.Value) + 
						Convert.ToDecimal(this.txtGNCDiciembre.Value);
				}
				#endregion Gastos Notas De Credito

				#region Notas De Credito Marketing
				if ((int)oDRConsulta["idTipo"] == 7)
				{
					if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtMNCEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
					if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtMNCFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
					if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtMNCMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
					if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtMNCAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
					if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtMNCMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
					if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtMNCJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
					if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtMNCJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
					if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtMNCAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
					if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtMNCSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
					if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtMNCOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
					if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtMNCNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
					if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtMNCDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

					this.txtMNCTotal.Value = 
						Convert.ToDecimal(this.txtMNCEnero.Value) + 
						Convert.ToDecimal(this.txtMNCFebrero.Value) + 
						Convert.ToDecimal(this.txtMNCMarzo.Value) + 
						Convert.ToDecimal(this.txtMNCAbril.Value) + 
						Convert.ToDecimal(this.txtMNCMayo.Value) + 
						Convert.ToDecimal(this.txtMNCJunio.Value) + 
						Convert.ToDecimal(this.txtMNCJulio.Value) + 
						Convert.ToDecimal(this.txtMNCAgosto.Value) + 
						Convert.ToDecimal(this.txtMNCSeptiembre.Value) + 
						Convert.ToDecimal(this.txtMNCOctubre.Value) + 
						Convert.ToDecimal(this.txtMNCNoviembre.Value) + 
						Convert.ToDecimal(this.txtMNCDiciembre.Value);
				}
				#endregion Notas De Credito Marketing

				#region Prestamos
				if ((int)oDRConsulta["idTipo"] == 8)
				{
					if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtPrestamosEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
					if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtPrestamosFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
					if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtPrestamosMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
					if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtPrestamosAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
					if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtPrestamosMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
					if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtPrestamosJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
					if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtPrestamosJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
					if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtPrestamosAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
					if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtPrestamosSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
					if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtPrestamosOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
					if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtPrestamosNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
					if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtPrestamosDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

					this.txtPrestamosTotal.Value = 
						Convert.ToDecimal(this.txtPrestamosEnero.Value) + 
						Convert.ToDecimal(this.txtPrestamosFebrero.Value) + 
						Convert.ToDecimal(this.txtPrestamosMarzo.Value) + 
						Convert.ToDecimal(this.txtPrestamosAbril.Value) + 
						Convert.ToDecimal(this.txtPrestamosMayo.Value) + 
						Convert.ToDecimal(this.txtPrestamosJunio.Value) + 
						Convert.ToDecimal(this.txtPrestamosJulio.Value) + 
						Convert.ToDecimal(this.txtPrestamosAgosto.Value) + 
						Convert.ToDecimal(this.txtPrestamosSeptiembre.Value) + 
						Convert.ToDecimal(this.txtPrestamosOctubre.Value) + 
						Convert.ToDecimal(this.txtPrestamosNoviembre.Value) + 
						Convert.ToDecimal(this.txtPrestamosDiciembre.Value);
				}
				#endregion Prestamos

				#region Fondo De Garantia
				if ((int)oDRConsulta["idTipo"] == 9)
				{
					if (oDRConsulta["Enero"] != System.DBNull.Value) this.txtFDGEnero.Value = Convert.ToDecimal(oDRConsulta["Enero"]);
					if (oDRConsulta["Febrero"] != System.DBNull.Value) this.txtFDGFebrero.Value = Convert.ToDecimal(oDRConsulta["Febrero"]);
					if (oDRConsulta["Marzo"] != System.DBNull.Value) this.txtFDGMarzo.Value = Convert.ToDecimal(oDRConsulta["Marzo"]);
					if (oDRConsulta["Abril"] != System.DBNull.Value) this.txtFDGAbril.Value = Convert.ToDecimal(oDRConsulta["Abril"]);
					if (oDRConsulta["Mayo"] != System.DBNull.Value) this.txtFDGMayo.Value = Convert.ToDecimal(oDRConsulta["Mayo"]);
					if (oDRConsulta["Junio"] != System.DBNull.Value) this.txtFDGJunio.Value = Convert.ToDecimal(oDRConsulta["Junio"]);
					if (oDRConsulta["Julio"] != System.DBNull.Value) this.txtFDGJulio.Value = Convert.ToDecimal(oDRConsulta["Julio"]);
					if (oDRConsulta["Agosto"] != System.DBNull.Value) this.txtFDGAgosto.Value = Convert.ToDecimal(oDRConsulta["Agosto"]);
					if (oDRConsulta["Septiembre"] != System.DBNull.Value) this.txtFDGSeptiembre.Value = Convert.ToDecimal(oDRConsulta["Septiembre"]);
					if (oDRConsulta["Octubre"] != System.DBNull.Value) this.txtFDGOctubre.Value = Convert.ToDecimal(oDRConsulta["Octubre"]);
					if (oDRConsulta["Noviembre"] != System.DBNull.Value) this.txtFDGNoviembre.Value = Convert.ToDecimal(oDRConsulta["Noviembre"]);
					if (oDRConsulta["Diciembre"] != System.DBNull.Value) this.txtFDGDiciembre.Value = Convert.ToDecimal(oDRConsulta["Diciembre"]);

					this.txtFDGTotal.Value = 
						Convert.ToDecimal(this.txtFDGEnero.Value) + 
						Convert.ToDecimal(this.txtFDGFebrero.Value) + 
						Convert.ToDecimal(this.txtFDGMarzo.Value) + 
						Convert.ToDecimal(this.txtFDGAbril.Value) + 
						Convert.ToDecimal(this.txtFDGMayo.Value) + 
						Convert.ToDecimal(this.txtFDGJunio.Value) + 
						Convert.ToDecimal(this.txtFDGJulio.Value) + 
						Convert.ToDecimal(this.txtFDGAgosto.Value) + 
						Convert.ToDecimal(this.txtFDGSeptiembre.Value) + 
						Convert.ToDecimal(this.txtFDGOctubre.Value) + 
						Convert.ToDecimal(this.txtFDGNoviembre.Value) + 
						Convert.ToDecimal(this.txtFDGDiciembre.Value);
				}
				#endregion Fondo De Garantia
			}
			#endregion Carga Datos

			#region Totales Ventas
			this.txtProyeccionVentaTotalEnero.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorEnero.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoEnero.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaEnero.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoEnero.Value);

			this.txtProyeccionVentaTotalFebrero.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorFebrero.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoFebrero.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaFebrero.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoFebrero.Value);

			this.txtProyeccionVentaTotalMarzo.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorMarzo.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoMarzo.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaMarzo.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoMarzo.Value);

			this.txtProyeccionVentaTotalAbril.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorAbril.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoAbril.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaAbril.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoAbril.Value);

			this.txtProyeccionVentaTotalMayo.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorMayo.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoMayo.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaMayo.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoMayo.Value);

			this.txtProyeccionVentaTotalJunio.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorJunio.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoJunio.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaJunio.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoJunio.Value);

			this.txtProyeccionVentaTotalJulio.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorJulio.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoJulio.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaJulio.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoJulio.Value);

			this.txtProyeccionVentaTotalAgosto.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorAgosto.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoAgosto.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaAgosto.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoAgosto.Value);

			this.txtProyeccionVentaTotalSeptiembre.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorSeptiembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoSeptiembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaSeptiembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoSeptiembre.Value);

			this.txtProyeccionVentaTotalOctubre.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorOctubre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoOctubre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaOctubre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoOctubre.Value);

			this.txtProyeccionVentaTotalNoviembre.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorNoviembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoNoviembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaNoviembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoNoviembre.Value);

			this.txtProyeccionVentaTotalDiciembre.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaDistribuidorDiciembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaEfectivoDiciembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaDiciembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaCreditoDiciembre.Value);

			this.txtProyeccionVentaTotal.Value = 
				Convert.ToDecimal(this.txtProyeccionVentaTotalEnero.Value) + 
				Convert.ToDecimal(this.txtProyeccionVentaTotalFebrero.Value) + 
				Convert.ToDecimal(this.txtProyeccionVentaTotalMarzo.Value) + 
				Convert.ToDecimal(this.txtProyeccionVentaTotalAbril.Value) + 
				Convert.ToDecimal(this.txtProyeccionVentaTotalMayo.Value) + 
				Convert.ToDecimal(this.txtProyeccionVentaTotalJunio.Value) + 
				Convert.ToDecimal(this.txtProyeccionVentaTotalJulio.Value) + 
				Convert.ToDecimal(this.txtProyeccionVentaTotalAgosto.Value) + 
				Convert.ToDecimal(this.txtProyeccionVentaTotalSeptiembre.Value) + 
				Convert.ToDecimal(this.txtProyeccionVentaTotalOctubre.Value) + 
				Convert.ToDecimal(this.txtProyeccionVentaTotalNoviembre.Value) + 
				Convert.ToDecimal(this.txtProyeccionVentaTotalDiciembre.Value);
			#endregion Totales Ventas

			#region Totales Margen
			this.txtProyeccionMargenTotalEnero.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorEnero.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoEnero.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaEnero.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoEnero.Value);

			this.txtProyeccionMargenTotalFebrero.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorFebrero.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoFebrero.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaFebrero.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoFebrero.Value);

			this.txtProyeccionMargenTotalMarzo.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorMarzo.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoMarzo.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaMarzo.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoMarzo.Value);

			this.txtProyeccionMargenTotalAbril.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorAbril.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoAbril.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaAbril.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoAbril.Value);

			this.txtProyeccionMargenTotalMayo.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorMayo.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoMayo.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaMayo.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoMayo.Value);

			this.txtProyeccionMargenTotalJunio.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorJunio.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoJunio.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaJunio.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoJunio.Value);

			this.txtProyeccionMargenTotalJulio.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorJulio.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoJulio.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaJulio.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoJulio.Value);

			this.txtProyeccionMargenTotalAgosto.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorAgosto.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoAgosto.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaAgosto.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoAgosto.Value);

			this.txtProyeccionMargenTotalSeptiembre.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorSeptiembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoSeptiembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaSeptiembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoSeptiembre.Value);

			this.txtProyeccionMargenTotalOctubre.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorOctubre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoOctubre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaOctubre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoOctubre.Value);

			this.txtProyeccionMargenTotalNoviembre.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorNoviembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoNoviembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaNoviembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoNoviembre.Value);

			this.txtProyeccionMargenTotalDiciembre.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenDistribuidorDiciembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenEfectivoDiciembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaDiciembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenCreditoDiciembre.Value);

			this.txtProyeccionMargenTotal.Value = 
				Convert.ToDecimal(this.txtProyeccionMargenTotalEnero.Value) + 
				Convert.ToDecimal(this.txtProyeccionMargenTotalFebrero.Value) + 
				Convert.ToDecimal(this.txtProyeccionMargenTotalMarzo.Value) + 
				Convert.ToDecimal(this.txtProyeccionMargenTotalAbril.Value) + 
				Convert.ToDecimal(this.txtProyeccionMargenTotalMayo.Value) + 
				Convert.ToDecimal(this.txtProyeccionMargenTotalJunio.Value) + 
				Convert.ToDecimal(this.txtProyeccionMargenTotalJulio.Value) + 
				Convert.ToDecimal(this.txtProyeccionMargenTotalAgosto.Value) + 
				Convert.ToDecimal(this.txtProyeccionMargenTotalSeptiembre.Value) + 
				Convert.ToDecimal(this.txtProyeccionMargenTotalOctubre.Value) + 
				Convert.ToDecimal(this.txtProyeccionMargenTotalNoviembre.Value) + 
				Convert.ToDecimal(this.txtProyeccionMargenTotalDiciembre.Value);
			#endregion Totales Margen

			#region Totales Ingresos
			this.txtTotalIngresosEnero.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalEnero.Value) + Convert.ToDecimal(this.txtCobranzaEnero.Value);
			this.txtTotalIngresosFebrero.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalFebrero.Value) + Convert.ToDecimal(this.txtCobranzaFebrero.Value);
			this.txtTotalIngresosMarzo.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalMarzo.Value) + Convert.ToDecimal(this.txtCobranzaMarzo.Value);
			this.txtTotalIngresosAbril.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalAbril.Value) + Convert.ToDecimal(this.txtCobranzaAbril.Value);
			this.txtTotalIngresosMayo.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalMayo.Value) + Convert.ToDecimal(this.txtCobranzaMayo.Value);
			this.txtTotalIngresosJunio.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalJunio.Value) + Convert.ToDecimal(this.txtCobranzaJunio.Value);
			this.txtTotalIngresosJulio.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalJulio.Value) + Convert.ToDecimal(this.txtCobranzaJulio.Value);
			this.txtTotalIngresosAgosto.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalAgosto.Value) + Convert.ToDecimal(this.txtCobranzaAgosto.Value);
			this.txtTotalIngresosSeptiembre.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalSeptiembre.Value) + Convert.ToDecimal(this.txtCobranzaSeptiembre.Value);
			this.txtTotalIngresosOctubre.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalOctubre.Value) + Convert.ToDecimal(this.txtCobranzaOctubre.Value);
			this.txtTotalIngresosNoviembre.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalNoviembre.Value) + Convert.ToDecimal(this.txtCobranzaNoviembre.Value);
			this.txtTotalIngresosDiciembre.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalDiciembre.Value) + Convert.ToDecimal(this.txtCobranzaDiciembre.Value);
			this.txtTotalIngresos.Value = Convert.ToDecimal(this.txtProyeccionMargenTotal.Value) + Convert.ToDecimal(this.txtCobranzaTotal.Value);
			#endregion Totales Ingresos

			#region Gastos
			sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualGastosAlmacenes '{0}', '{1}', {2}, {3}, {4}", 
				dtFechaInicial.ToString("yyyyMMdd"), dtFechaFinal.ToString("yyyyMMdd"), iPeriodo, IBodega, (bool)this.optContrato.Value);
			this.uGridGastos.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			#region Suma de Gastos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)				
			{
				if (drG.Cells["ENERO"].Value != System.DBNull.Value) this.txtTotalGastosEnero.Value = Convert.ToDecimal(this.txtTotalGastosEnero.Value) + Convert.ToDecimal(drG.Cells["ENERO"].Value);
				if (drG.Cells["FEBRERO"].Value != System.DBNull.Value) this.txtTotalGastosFebrero.Value = Convert.ToDecimal(this.txtTotalGastosFebrero.Value) + Convert.ToDecimal(drG.Cells["FEBRERO"].Value);
				if (drG.Cells["MARZO"].Value != System.DBNull.Value) this.txtTotalGastosMarzo.Value = Convert.ToDecimal(this.txtTotalGastosMarzo.Value) + Convert.ToDecimal(drG.Cells["MARZO"].Value);
				if (drG.Cells["ABRIL"].Value != System.DBNull.Value) this.txtTotalGastosAbril.Value = Convert.ToDecimal(this.txtTotalGastosAbril.Value) + Convert.ToDecimal(drG.Cells["ABRIL"].Value);
				if (drG.Cells["MAYO"].Value != System.DBNull.Value) this.txtTotalGastosMayo.Value = Convert.ToDecimal(this.txtTotalGastosMayo.Value) + Convert.ToDecimal(drG.Cells["MAYO"].Value);
				if (drG.Cells["JUNIO"].Value != System.DBNull.Value) this.txtTotalGastosJunio.Value = Convert.ToDecimal(this.txtTotalGastosJunio.Value) + Convert.ToDecimal(drG.Cells["JUNIO"].Value);
				if (drG.Cells["JULIO"].Value != System.DBNull.Value) this.txtTotalGastosJulio.Value = Convert.ToDecimal(this.txtTotalGastosJulio.Value) + Convert.ToDecimal(drG.Cells["JULIO"].Value);
				if (drG.Cells["AGOSTO"].Value != System.DBNull.Value) this.txtTotalGastosAgosto.Value = Convert.ToDecimal(this.txtTotalGastosAgosto.Value) + Convert.ToDecimal(drG.Cells["AGOSTO"].Value);
				if (drG.Cells["SEPTIEMBRE"].Value != System.DBNull.Value) this.txtTotalGastosSeptiembre.Value = Convert.ToDecimal(this.txtTotalGastosSeptiembre.Value) + Convert.ToDecimal(drG.Cells["SEPTIEMBRE"].Value);
				if (drG.Cells["OCTUBRE"].Value != System.DBNull.Value) this.txtTotalGastosOctubre.Value = Convert.ToDecimal(this.txtTotalGastosOctubre.Value) + Convert.ToDecimal(drG.Cells["OCTUBRE"].Value);
				if (drG.Cells["NOVIEMBRE"].Value != System.DBNull.Value) this.txtTotalGastosNoviembre.Value = Convert.ToDecimal(this.txtTotalGastosNoviembre.Value) + Convert.ToDecimal(drG.Cells["NOVIEMBRE"].Value);
				if (drG.Cells["DICIEMBRE"].Value != System.DBNull.Value) this.txtTotalGastosDiciembre.Value = Convert.ToDecimal(this.txtTotalGastosDiciembre.Value) + Convert.ToDecimal(drG.Cells["DICIEMBRE"].Value);
			}

			this.txtTotalGastos.Value = 
				Convert.ToDecimal(this.txtTotalGastosEnero.Value) + 
				Convert.ToDecimal(this.txtTotalGastosFebrero.Value) + 
				Convert.ToDecimal(this.txtTotalGastosMarzo.Value) + 
				Convert.ToDecimal(this.txtTotalGastosAbril.Value) + 
				Convert.ToDecimal(this.txtTotalGastosMayo.Value) + 
				Convert.ToDecimal(this.txtTotalGastosJunio.Value) + 
				Convert.ToDecimal(this.txtTotalGastosJulio.Value) + 
				Convert.ToDecimal(this.txtTotalGastosAgosto.Value) + 
				Convert.ToDecimal(this.txtTotalGastosSeptiembre.Value) + 
				Convert.ToDecimal(this.txtTotalGastosOctubre.Value) + 
				Convert.ToDecimal(this.txtTotalGastosNoviembre.Value) + 
				Convert.ToDecimal(this.txtTotalGastosDiciembre.Value);
			#endregion Suma de Gastos
			#endregion Gastos

			#region Utilidad Neta
			this.txtResultadoEnero.Value = 
				Convert.ToDecimal(this.txtTotalIngresosEnero.Value) - 
				Convert.ToDecimal(this.txtTotalGastosEnero.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroEnero.Value) +
				Convert.ToDecimal(this.txtGNCEnero.Value) +
				Convert.ToDecimal(this.txtMNCEnero.Value);

			this.txtResultadoFebrero.Value = 
				Convert.ToDecimal(this.txtTotalIngresosFebrero.Value) - 
				Convert.ToDecimal(this.txtTotalGastosFebrero.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroFebrero.Value) +
				Convert.ToDecimal(this.txtGNCFebrero.Value) +
				Convert.ToDecimal(this.txtMNCFebrero.Value);

			this.txtResultadoMarzo.Value = 
				Convert.ToDecimal(this.txtTotalIngresosMarzo.Value) - 
				Convert.ToDecimal(this.txtTotalGastosMarzo.Value) -
				Convert.ToDecimal(this.txtGastosRetiroMarzo.Value) +
				Convert.ToDecimal(this.txtGNCMarzo.Value) +
				Convert.ToDecimal(this.txtMNCMarzo.Value);

			this.txtResultadoAbril.Value = 
				Convert.ToDecimal(this.txtTotalIngresosAbril.Value) - 
				Convert.ToDecimal(this.txtTotalGastosAbril.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroAbril.Value) +
				Convert.ToDecimal(this.txtGNCAbril.Value) +
				Convert.ToDecimal(this.txtMNCAbril.Value);

			this.txtResultadoMayo.Value = 
				Convert.ToDecimal(this.txtTotalIngresosMayo.Value) - 
				Convert.ToDecimal(this.txtTotalGastosMayo.Value) -
				Convert.ToDecimal(this.txtGastosRetiroMayo.Value) +
				Convert.ToDecimal(this.txtGNCMayo.Value) +
				Convert.ToDecimal(this.txtMNCMayo.Value);

			this.txtResultadoJunio.Value = 
				Convert.ToDecimal(this.txtTotalIngresosJunio.Value) - 
				Convert.ToDecimal(this.txtTotalGastosJunio.Value) -
				Convert.ToDecimal(this.txtGastosRetiroJunio.Value) +
				Convert.ToDecimal(this.txtGNCJunio.Value) +
				Convert.ToDecimal(this.txtMNCJunio.Value);

			this.txtResultadoJulio.Value = 
				Convert.ToDecimal(this.txtTotalIngresosJulio.Value) - 
				Convert.ToDecimal(this.txtTotalGastosJulio.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroJulio.Value) +
				Convert.ToDecimal(this.txtGNCJulio.Value) +
				Convert.ToDecimal(this.txtMNCJulio.Value);

			this.txtResultadoAgosto.Value = 
				Convert.ToDecimal(this.txtTotalIngresosAgosto.Value) - 
				Convert.ToDecimal(this.txtTotalGastosAgosto.Value) -
				Convert.ToDecimal(this.txtGastosRetiroAgosto.Value) +
				Convert.ToDecimal(this.txtGNCAgosto.Value) +
				Convert.ToDecimal(this.txtMNCAgosto.Value);

			this.txtResultadoSeptiembre.Value = 
				Convert.ToDecimal(this.txtTotalIngresosSeptiembre.Value) - 
				Convert.ToDecimal(this.txtTotalGastosSeptiembre.Value) -
				Convert.ToDecimal(this.txtGastosRetiroSeptiembre.Value) +
				Convert.ToDecimal(this.txtGNCSeptiembre.Value) +
				Convert.ToDecimal(this.txtMNCSeptiembre.Value);

			this.txtResultadoOctubre.Value = 
				Convert.ToDecimal(this.txtTotalIngresosOctubre.Value) - 
				Convert.ToDecimal(this.txtTotalGastosOctubre.Value) -
				Convert.ToDecimal(this.txtGastosRetiroOctubre.Value) +
				Convert.ToDecimal(this.txtGNCOctubre.Value) +
				Convert.ToDecimal(this.txtMNCOctubre.Value);

			this.txtResultadoNoviembre.Value = 
				Convert.ToDecimal(this.txtTotalIngresosNoviembre.Value) - 
				Convert.ToDecimal(this.txtTotalGastosNoviembre.Value) -
				Convert.ToDecimal(this.txtGastosRetiroNoviembre.Value) +
				Convert.ToDecimal(this.txtGNCNoviembre.Value) +
				Convert.ToDecimal(this.txtMNCNoviembre.Value);

			this.txtResultadoDiciembre.Value = 
				Convert.ToDecimal(this.txtTotalIngresosDiciembre.Value) - 
				Convert.ToDecimal(this.txtTotalGastosDiciembre.Value) -
				Convert.ToDecimal(this.txtGastosRetiroDiciembre.Value) +
				Convert.ToDecimal(this.txtGNCDiciembre.Value) +
				Convert.ToDecimal(this.txtMNCDiciembre.Value);

			this.txtResultadoTotal.Value = 
				Convert.ToDecimal(this.txtTotalIngresos.Value) - 
				Convert.ToDecimal(this.txtTotalGastos.Value) -
				Convert.ToDecimal(this.txtGastosRetiroTotal.Value) +
				Convert.ToDecimal(this.txtGNCTotal.Value) +
				Convert.ToDecimal(this.txtMNCTotal.Value);
			#endregion Utilidad Neta

			#region Total General
			this.txtTotalGeneralEnero.Value = 
				Convert.ToDecimal(this.txtResultadoEnero.Value) - 
				Convert.ToDecimal(this.txtPrestamosEnero.Value) + 
				Convert.ToDecimal(this.txtFDGEnero.Value);

			this.txtTotalGeneralFebrero.Value = 
				Convert.ToDecimal(this.txtResultadoFebrero.Value) - 
				Convert.ToDecimal(this.txtPrestamosFebrero.Value) + 
				Convert.ToDecimal(this.txtFDGFebrero.Value);

			this.txtTotalGeneralMarzo.Value = 
				Convert.ToDecimal(this.txtResultadoMarzo.Value) - 
				Convert.ToDecimal(this.txtPrestamosMarzo.Value) + 
				Convert.ToDecimal(this.txtFDGMarzo.Value);

			this.txtTotalGeneralAbril.Value = 
				Convert.ToDecimal(this.txtResultadoAbril.Value) - 
				Convert.ToDecimal(this.txtPrestamosAbril.Value) + 
				Convert.ToDecimal(this.txtFDGAbril.Value);

			this.txtTotalGeneralMayo.Value = 
				Convert.ToDecimal(this.txtResultadoMayo.Value) - 
				Convert.ToDecimal(this.txtPrestamosMayo.Value) +
				Convert.ToDecimal(this.txtFDGMayo.Value);

			this.txtTotalGeneralJunio.Value = 
				Convert.ToDecimal(this.txtResultadoJunio.Value) - 
				Convert.ToDecimal(this.txtPrestamosJunio.Value) + 
				Convert.ToDecimal(this.txtFDGJunio.Value);

			this.txtTotalGeneralJulio.Value = 
				Convert.ToDecimal(this.txtResultadoJulio.Value) - 
				Convert.ToDecimal(this.txtPrestamosJulio.Value) + 
				Convert.ToDecimal(this.txtFDGJulio.Value);

			this.txtTotalGeneralAgosto.Value = 
				Convert.ToDecimal(this.txtResultadoAgosto.Value) - 
				Convert.ToDecimal(this.txtPrestamosAgosto.Value) + 
				Convert.ToDecimal(this.txtFDGAgosto.Value);

			this.txtTotalGeneralSeptiembre.Value = 
				Convert.ToDecimal(this.txtResultadoSeptiembre.Value) - 
				Convert.ToDecimal(this.txtPrestamosSeptiembre.Value) +
				Convert.ToDecimal(this.txtFDGSeptiembre.Value);

			this.txtTotalGeneralOctubre.Value = 
				Convert.ToDecimal(this.txtResultadoOctubre.Value) - 
				Convert.ToDecimal(this.txtPrestamosOctubre.Value) +
				Convert.ToDecimal(this.txtFDGOctubre.Value);

			this.txtTotalGeneralNoviembre.Value = 
				Convert.ToDecimal(this.txtResultadoNoviembre.Value) - 
				Convert.ToDecimal(this.txtPrestamosNoviembre.Value) + 
				Convert.ToDecimal(this.txtFDGNoviembre.Value);

			this.txtTotalGeneralDiciembre.Value = 
				Convert.ToDecimal(this.txtResultadoDiciembre.Value) - 
				Convert.ToDecimal(this.txtPrestamosDiciembre.Value) + 
				Convert.ToDecimal(this.txtFDGDiciembre.Value);

			this.txtTotalGeneral.Value = 
				Convert.ToDecimal(this.txtResultadoTotal.Value) - 
				Convert.ToDecimal(this.txtPrestamosTotal.Value) +
				Convert.ToDecimal(this.txtFDGTotal.Value);
			#endregion Total General
		}

		private void frmProyeccionPresupuestoAlmacen_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200209");

			if (!Funcion.Permiso("927", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Presupuesto por Almacen", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
        if (miAcceso.BExportar) this.btnExcel.Enabled = true;
			this.txtAño.MaxValue = DateTime.Today.Year;
			this.txtAño.Value = DateTime.Today.Year;
//			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre, Codigo From Bodega Where Activo = 1 And Factura = 1 And Caja = 1 Order By Nombre");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre, Codigo From Bodega Where Activo = 1 And Factura = 1 And Caja = 1 AND (ISNULL (DATEPART(Year, FechaCierre),'2030') )> '" + Convert.ToDateTime(dtFechaFinal).Year.ToString()+ "' Order By Nombre");
	      
		}

		private void DatosEnero()
		{
			DateTime dtDesde = DateTime.Parse("01/01/" + iPeriodo.ToString());
			DateTime dtHasta = dtDesde;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionEneroVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorEnero.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoEnero.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaEnero.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoEnero.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorEnero.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoEnero.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaEnero.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoEnero.Value = dr.GetDecimal(8);
				this.txtCobranzaEnero.Value = dr.GetDecimal(9);
				this.txtGastosRetiroEnero.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalEnero.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorEnero.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoEnero.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaEnero.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoEnero.Value);

			this.txtProyeccionMargenTotalEnero.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorEnero.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoEnero.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaEnero.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoEnero.Value);

			this.txtTotalIngresosEnero.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalEnero.Value) + Convert.ToDecimal(this.txtCobranzaEnero.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)				
				if (drG.Cells["ENERO"].Value != System.DBNull.Value)
					this.txtTotalGastosEnero.Value = Convert.ToDecimal(this.txtTotalGastosEnero.Value) + Convert.ToDecimal(drG.Cells["ENERO"].Value);			
			
			this.txtResultadoEnero.Value = Convert.ToDecimal(this.txtTotalIngresosEnero.Value) - Convert.ToDecimal(this.txtTotalGastosEnero.Value) - Convert.ToDecimal(this.txtGastosRetiroEnero.Value);			

			//this.txtSMAGastosRetirosFebrero.Value = Convert.ToDecimal(this.txtTotalGastosEnero.Value) + Convert.ToDecimal(this.txtGastosRetiroEnero.Value);
		}

		private void DatosFebrero()
		{
			DateTime dtDesde = DateTime.Parse("01/02/" + iPeriodo.ToString());
			DateTime dtHasta = dtDesde;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionFebreroVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorFebrero.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoFebrero.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaFebrero.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoFebrero.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorFebrero.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoFebrero.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaFebrero.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoFebrero.Value = dr.GetDecimal(8);
				this.txtCobranzaFebrero.Value = dr.GetDecimal(9);
				this.txtGastosRetiroFebrero.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalFebrero.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorFebrero.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoFebrero.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaFebrero.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoFebrero.Value);

			this.txtProyeccionMargenTotalFebrero.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorFebrero.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoFebrero.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaFebrero.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoFebrero.Value);

			this.txtTotalIngresosFebrero.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalFebrero.Value) + Convert.ToDecimal(this.txtCobranzaFebrero.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)
			{	
				if (drG.Cells["FEBRERO"].Value != System.DBNull.Value)
					this.txtTotalGastosFebrero.Value = Convert.ToDecimal(this.txtTotalGastosFebrero.Value) + Convert.ToDecimal(drG.Cells["FEBRERO"].Value);
			}
			
			this.txtResultadoFebrero.Value = Convert.ToDecimal(this.txtTotalIngresosFebrero.Value) - Convert.ToDecimal(this.txtTotalGastosFebrero.Value) - Convert.ToDecimal(this.txtGastosRetiroFebrero.Value);						
		}

		private void DatosMarzo()
		{
			DateTime dtDesde = DateTime.Parse("01/03/" + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Today;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}',{2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionMarzoVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorMarzo.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoMarzo.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaMarzo.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoMarzo.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorMarzo.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoMarzo.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaMarzo.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoMarzo.Value = dr.GetDecimal(8);
				this.txtCobranzaMarzo.Value = dr.GetDecimal(9);
				this.txtGastosRetiroMarzo.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalMarzo.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorMarzo.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoMarzo.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaMarzo.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoMarzo.Value);

			this.txtProyeccionMargenTotalMarzo.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorMarzo.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoMarzo.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaMarzo.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoMarzo.Value);

			this.txtTotalIngresosMarzo.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalMarzo.Value) + Convert.ToDecimal(this.txtCobranzaMarzo.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)
			{	
				if (drG.Cells["MARZO"].Value != System.DBNull.Value)
					this.txtTotalGastosMarzo.Value = Convert.ToDecimal(this.txtTotalGastosMarzo.Value) + Convert.ToDecimal(drG.Cells["MARZO"].Value);
			}
			
			this.txtResultadoMarzo.Value = Convert.ToDecimal(this.txtTotalIngresosMarzo.Value) - 
				Convert.ToDecimal(this.txtTotalGastosMarzo.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroMarzo.Value);
		}

		private void DatosAbril()
		{
			DateTime dtDesde = DateTime.Parse("01/04/" + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Today;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionAbrilVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorAbril.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoAbril.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaAbril.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoAbril.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorAbril.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoAbril.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaAbril.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoAbril.Value = dr.GetDecimal(8);
				this.txtCobranzaAbril.Value = dr.GetDecimal(9);
				this.txtGastosRetiroAbril.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalAbril.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorAbril.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoAbril.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaAbril.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoAbril.Value);

			this.txtProyeccionMargenTotalAbril.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorAbril.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoAbril.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaAbril.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoAbril.Value);

			this.txtTotalIngresosAbril.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalAbril.Value) + Convert.ToDecimal(this.txtCobranzaAbril.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)
			{	
				if (drG.Cells["ABRIL"].Value != System.DBNull.Value)
					this.txtTotalGastosAbril.Value = Convert.ToDecimal(this.txtTotalGastosAbril.Value) + Convert.ToDecimal(drG.Cells["ABRIL"].Value);
			}
			
			this.txtResultadoAbril.Value = Convert.ToDecimal(this.txtTotalIngresosAbril.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroAbril.Value) - 
				Convert.ToDecimal(this.txtTotalGastosAbril.Value);			
		}

		private void DatosMayo()
		{
			DateTime dtDesde = DateTime.Parse("01/05/" + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Today;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionMayoVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorMayo.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoMayo.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaMayo.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoMayo.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorMayo.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoMayo.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaMayo.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoMayo.Value = dr.GetDecimal(8);
				this.txtCobranzaMayo.Value = dr.GetDecimal(9);
				this.txtGastosRetiroMayo.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalMayo.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorMayo.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoMayo.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaMayo.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoMayo.Value);

			this.txtProyeccionMargenTotalMayo.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorMayo.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoMayo.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaMayo.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoMayo.Value);

			this.txtTotalIngresosMayo.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalMayo.Value) + Convert.ToDecimal(this.txtCobranzaMayo.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)
			{	
				if (drG.Cells["Mayo"].Value != System.DBNull.Value)
					this.txtTotalGastosMayo.Value = Convert.ToDecimal(this.txtTotalGastosMayo.Value) + Convert.ToDecimal(drG.Cells["Mayo"].Value);
			}

			this.txtResultadoMayo.Value = Convert.ToDecimal(this.txtTotalIngresosMayo.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroMayo.Value) -																		
				Convert.ToDecimal(this.txtTotalGastosMayo.Value);
		}

		private void DatosJunio()
		{
			DateTime dtDesde = DateTime.Parse("01/06/" + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Today;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionJunioVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorJunio.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoJunio.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaJunio.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoJunio.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorJunio.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoJunio.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaJunio.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoJunio.Value = dr.GetDecimal(8);
				this.txtCobranzaJunio.Value = dr.GetDecimal(9);
				this.txtGastosRetiroJunio.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalJunio.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorJunio.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoJunio.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaJunio.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoJunio.Value);

			this.txtProyeccionMargenTotalJunio.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorJunio.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoJunio.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaJunio.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoJunio.Value);

			this.txtTotalIngresosJunio.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalJunio.Value) + Convert.ToDecimal(this.txtCobranzaJunio.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)
			{	
				if (drG.Cells["Junio"].Value != System.DBNull.Value)
					this.txtTotalGastosJunio.Value = Convert.ToDecimal(this.txtTotalGastosJunio.Value) + Convert.ToDecimal(drG.Cells["Junio"].Value);
			}
			
			this.txtResultadoJunio.Value = Convert.ToDecimal(this.txtTotalIngresosJunio.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroJunio.Value) -																		
				Convert.ToDecimal(this.txtTotalGastosJunio.Value);
		}

		private void DatosJulio()
		{
			DateTime dtDesde = DateTime.Parse("01/07/" + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Today;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionJulioVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorJulio.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoJulio.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaJulio.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoJulio.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorJulio.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoJulio.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaJulio.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoJulio.Value = dr.GetDecimal(8);
				this.txtCobranzaJulio.Value = dr.GetDecimal(9);
				this.txtGastosRetiroJulio.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalJulio.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorJulio.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoJulio.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaJulio.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoJulio.Value);

			this.txtProyeccionMargenTotalJulio.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorJulio.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoJulio.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaJulio.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoJulio.Value);

			this.txtTotalIngresosJulio.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalJulio.Value) + Convert.ToDecimal(this.txtCobranzaJulio.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)
			{	
				if (drG.Cells["Julio"].Value != System.DBNull.Value)
					this.txtTotalGastosJulio.Value = Convert.ToDecimal(this.txtTotalGastosJulio.Value) + Convert.ToDecimal(drG.Cells["Julio"].Value);
			}
			
			this.txtResultadoJulio.Value = Convert.ToDecimal(this.txtTotalIngresosJulio.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroJulio.Value) -																		
				Convert.ToDecimal(this.txtTotalGastosJulio.Value);
		}

		private void DatosAgosto()
		{
			DateTime dtDesde = DateTime.Parse("01/08/" + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Today;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionAgostoVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorAgosto.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoAgosto.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaAgosto.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoAgosto.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorAgosto.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoAgosto.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaAgosto.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoAgosto.Value = dr.GetDecimal(8);
				this.txtCobranzaAgosto.Value = dr.GetDecimal(9);
				this.txtGastosRetiroAgosto.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalAgosto.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorAgosto.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoAgosto.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaAgosto.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoAgosto.Value);

			this.txtProyeccionMargenTotalAgosto.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorAgosto.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoAgosto.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaAgosto.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoAgosto.Value);

			this.txtTotalIngresosAgosto.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalAgosto.Value) + Convert.ToDecimal(this.txtCobranzaAgosto.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)
			{	
				if (drG.Cells["Agosto"].Value != System.DBNull.Value)
					this.txtTotalGastosAgosto.Value = Convert.ToDecimal(this.txtTotalGastosAgosto.Value) + Convert.ToDecimal(drG.Cells["Agosto"].Value);
			}
			
			this.txtResultadoAgosto.Value = Convert.ToDecimal(this.txtTotalIngresosAgosto.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroAgosto.Value) -																		
				Convert.ToDecimal(this.txtTotalGastosAgosto.Value);	
		}

		private void DatosSeptiembre()
		{
			DateTime dtDesde = DateTime.Parse("01/09/" + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Today;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionSeptiembreVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorSeptiembre.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoSeptiembre.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaSeptiembre.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoSeptiembre.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorSeptiembre.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoSeptiembre.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaSeptiembre.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoSeptiembre.Value = dr.GetDecimal(8);
				this.txtCobranzaSeptiembre.Value = dr.GetDecimal(9);
				this.txtGastosRetiroSeptiembre.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalSeptiembre.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorSeptiembre.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoSeptiembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaSeptiembre.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoSeptiembre.Value);

			this.txtProyeccionMargenTotalSeptiembre.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorSeptiembre.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoSeptiembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaSeptiembre.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoSeptiembre.Value);

			this.txtTotalIngresosSeptiembre.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalSeptiembre.Value) + Convert.ToDecimal(this.txtCobranzaSeptiembre.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)
			{	
				if (drG.Cells["Septiembre"].Value != System.DBNull.Value)
					this.txtTotalGastosSeptiembre.Value = Convert.ToDecimal(this.txtTotalGastosSeptiembre.Value) + Convert.ToDecimal(drG.Cells["Septiembre"].Value);
			}
			
			//this.txtSMAGastosRetirosSeptiembre.Value = Convert.ToDecimal(this.txtTotalGastosJulio.Value) + Convert.ToDecimal(this.txtGastosRetiroJulio.Value);

			this.txtResultadoSeptiembre.Value = Convert.ToDecimal(this.txtTotalIngresosSeptiembre.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroSeptiembre.Value) -																		
				Convert.ToDecimal(this.txtTotalGastosSeptiembre.Value);		
		}

		private void DatosOctubre()
		{
			DateTime dtDesde = DateTime.Parse("01/10/" + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Today;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionOctubreVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorOctubre.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoOctubre.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaOctubre.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoOctubre.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorOctubre.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoOctubre.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaOctubre.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoOctubre.Value = dr.GetDecimal(8);
				this.txtCobranzaOctubre.Value = dr.GetDecimal(9);
				this.txtGastosRetiroOctubre.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalOctubre.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorOctubre.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoOctubre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaOctubre.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoOctubre.Value);

			this.txtProyeccionMargenTotalOctubre.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorOctubre.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoOctubre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaOctubre.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoOctubre.Value);

			this.txtTotalIngresosOctubre.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalOctubre.Value) + Convert.ToDecimal(this.txtCobranzaOctubre.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)
			{	
				if (drG.Cells["Octubre"].Value != System.DBNull.Value)
					this.txtTotalGastosOctubre.Value = Convert.ToDecimal(this.txtTotalGastosOctubre.Value) + Convert.ToDecimal(drG.Cells["Octubre"].Value);
			}
			
			this.txtResultadoOctubre.Value = Convert.ToDecimal(this.txtTotalIngresosOctubre.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroOctubre.Value) -																		
				Convert.ToDecimal(this.txtTotalGastosOctubre.Value);
		}

		private void DatosNoviembre()
		{
			DateTime dtDesde = DateTime.Parse("01/11/" + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Today;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionNoviembreVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorNoviembre.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoNoviembre.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaNoviembre.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoNoviembre.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorNoviembre.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoNoviembre.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaNoviembre.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoNoviembre.Value = dr.GetDecimal(8);
				this.txtCobranzaNoviembre.Value = dr.GetDecimal(9);
				this.txtGastosRetiroNoviembre.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalNoviembre.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorNoviembre.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoNoviembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaNoviembre.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoNoviembre.Value);

			this.txtProyeccionMargenTotalNoviembre.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorNoviembre.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoNoviembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaNoviembre.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoNoviembre.Value);

			this.txtTotalIngresosNoviembre.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalNoviembre.Value) + Convert.ToDecimal(this.txtCobranzaNoviembre.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)
			{	
				if (drG.Cells["Noviembre"].Value != System.DBNull.Value)
					this.txtTotalGastosNoviembre.Value = Convert.ToDecimal(this.txtTotalGastosNoviembre.Value) + Convert.ToDecimal(drG.Cells["Noviembre"].Value);
			}
			
			//this.txtSMAGastosRetirosNoviembre.Value = Convert.ToDecimal(this.txtTotalGastosJulio.Value) + Convert.ToDecimal(this.txtGastosRetiroJulio.Value);

			this.txtResultadoNoviembre.Value = Convert.ToDecimal(this.txtTotalIngresosNoviembre.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroNoviembre.Value) -																		
				Convert.ToDecimal(this.txtTotalGastosNoviembre.Value);			
		}

		private void DatosDiciembre()
		{
			DateTime dtDesde = DateTime.Parse("01/12/" + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Today;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			string sSQL = string.Format("Exec ProyeccionRealidadPresupuestoAnualPorAlmacen '{0}', '{1}', {2}", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int) this.cmbBodega.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtProyeccionDiciembreVenta.Value = dr.GetDecimal(0);
				this.txtProyeccionVentaDistribuidorDiciembre.Value = dr.GetDecimal(1);
				this.txtProyeccionVentaEfectivoDiciembre.Value = dr.GetDecimal(2);
				this.txtProyeccionVentaTarjetaDiciembre.Value = dr.GetDecimal(3);
				this.txtProyeccionVentaCreditoDiciembre.Value = dr.GetDecimal(4);
				
				this.txtProyeccionMargenDistribuidorDiciembre.Value = dr.GetDecimal(5);
				this.txtProyeccionMargenEfectivoDiciembre.Value = dr.GetDecimal(6);
				this.txtProyeccionMargenTarjetaDiciembre.Value = dr.GetDecimal(7);
				this.txtProyeccionMargenCreditoDiciembre.Value = dr.GetDecimal(8);
				this.txtCobranzaDiciembre.Value = dr.GetDecimal(9);
				this.txtGastosRetiroDiciembre.Value = dr.GetDecimal(10);
			}
			dr.Close();

			this.txtProyeccionVentaTotalDiciembre.Value = Convert.ToDecimal(this.txtProyeccionVentaDistribuidorDiciembre.Value) + Convert.ToDecimal(this.txtProyeccionVentaEfectivoDiciembre.Value) +
				Convert.ToDecimal(this.txtProyeccionVentaTarjetaDiciembre.Value) + Convert.ToDecimal(this.txtProyeccionVentaCreditoDiciembre.Value);

			this.txtProyeccionMargenTotalDiciembre.Value = Convert.ToDecimal(this.txtProyeccionMargenDistribuidorDiciembre.Value) + Convert.ToDecimal(this.txtProyeccionMargenEfectivoDiciembre.Value) +
				Convert.ToDecimal(this.txtProyeccionMargenTarjetaDiciembre.Value) + Convert.ToDecimal(this.txtProyeccionMargenCreditoDiciembre.Value);

			this.txtTotalIngresosDiciembre.Value = Convert.ToDecimal(this.txtProyeccionMargenTotalDiciembre.Value) + Convert.ToDecimal(this.txtCobranzaDiciembre.Value);
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drG in this.uGridGastos.Rows.All)
			{	
				if (drG.Cells["Diciembre"].Value != System.DBNull.Value)
					this.txtTotalGastosDiciembre.Value = Convert.ToDecimal(this.txtTotalGastosDiciembre.Value) + Convert.ToDecimal(drG.Cells["Diciembre"].Value);
			}
			
			//this.txtSMAGastosRetirosDiciembre.Value = Convert.ToDecimal(this.txtGastosRetiroNoviembre.Value); //Convert.ToDecimal(this.txtTotalGastosNoviembre.Value) + 

			this.txtResultadoDiciembre.Value = Convert.ToDecimal(this.txtTotalIngresosDiciembre.Value) - 
				Convert.ToDecimal(this.txtGastosRetiroDiciembre.Value) -																		
				Convert.ToDecimal(this.txtTotalGastosDiciembre.Value);			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmProyeccionPresupuestoAlmacen_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
//			if (e.KeyCode == Keys.F5) this.ConsultaGeneral();
		}

		private void uGridGastos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.ToString();
			string sPeriodo = sColumna.Substring(3).ToUpper();
			
			if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay gastos registrado en el periodo {0}", sColumna), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.MostrarDetalle(e.Cell.Column.ToString(), (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, sPeriodo, e.Cell.Row.Cells["Grupo"].Value.ToString());
		
		}

		private void MostrarDetalle(string sColumna, int idCuentaAgrupa, string sPeriodo, string sGasto)
		{			
			string sMes = "01/";
			if (sColumna == "btnFebrero") sMes = "02/";
			if (sColumna == "btnMarzo") sMes = "03/";
			if (sColumna == "btnAbril") sMes = "04/";
			if (sColumna == "btnMayo") sMes = "05/";
			if (sColumna == "btnJunio") sMes = "06/";
			if (sColumna == "btnJulio") sMes = "07/";
			if (sColumna == "btnAgosto") sMes = "08/";
			if (sColumna == "btnSeptiembre") sMes = "09/";
			if (sColumna == "btnOctubre") sMes = "10/";
			if (sColumna == "btnNoviembre") sMes = "11/";
			if (sColumna == "btnDiciembre") sMes = "12/";
			
			DateTime dtDesde = DateTime.Parse("01/" + sMes + iPeriodo.ToString());
			DateTime dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			if ((int)this.uGridGastos.ActiveRow.Cells["idCuentaAgrupa"].Value == 1)
			{
				 frmRolReporteDetalleIngresosPorDepartamento rdind = new frmRolReporteDetalleIngresosPorDepartamento((int)this.cmbBodega.Value, dtHasta, false) ;
				rdind.ShowDialog();
			}
			else
			{
					frmPresupuestoResumenGastos prg = new frmPresupuestoResumenGastos(dtDesde, dtHasta, idCuentaAgrupa, sPeriodo, sGasto,(int)this.cmbBodega.Value, false, (int)this.txtAño.Value);
					prg.ShowDialog();
			}
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			this.ConsultaGeneral();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridGastos);
		}
		private void AbreDocumentos(int iMes, int iAño, int iOrigen)
		{
//			frmListaProyeccionDocumentos LPD = new frmListaProyeccionDocumentos(iMes, iAño, iOrigen);
//			LPD.ShowDialog();
		}
		private void vaciar ()
		{
			this.	txtTotalGeneral	.Value = 	0	;
			this.	 txtTotalGeneralNoviembre	.Value = 	0	;
			this.	txtTotalGeneralDiciembre	.Value = 	0	;
			this.	 txtTotalGeneralSeptiembre	.Value = 	0	;
			this.	 txtTotalGeneralOctubre	.Value = 	0	;
			this.	 txtTotalGeneralJulio	.Value = 	0	;
			this.	 txtTotalGeneralAgosto	.Value = 	0	;
			this.	 txtTotalGeneralMayo	.Value = 	0	;
			this.	 txtTotalGeneralJunio	.Value = 	0	;
			this.	 txtTotalGeneralMarzo	.Value = 	0	;
			this.	txtTotalGeneralAbril	.Value = 	0	;
			this.	 txtTotalGeneralFebrero	.Value = 	0	;
			this.	 txtTotalGeneralEnero	.Value = 	0	;
			this.	 txtFDGTotal	.Value = 	0	;
			this.	txtFDGNoviembre	.Value = 	0	;
			this.	 txtFDGDiciembre	.Value = 	0	;
			this.	 txtFDGSeptiembre	.Value = 	0	;
			this.	 txtFDGOctubre	.Value = 	0	;
			this.	txtFDGJulio	.Value = 	0	;
			this.	 txtFDGAgosto	.Value = 	0	;
			this.	 txtFDGMayo	.Value = 	0	;
			this.	 txtFDGJunio	.Value = 	0	;
			this.	 txtFDGMarzo	.Value = 	0	;
			this.	 txtFDGAbril	.Value = 	0	;
			this.	txtFDGFebrero	.Value = 	0	;
			this.	 txtFDGEnero	.Value = 	0	;
			this.	 txtPrestamosTotal	.Value = 	0	;
			this.	txtPrestamosNoviembre	.Value = 	0	;
			this.	 txtPrestamosDiciembre	.Value = 	0	;
			this.	 txtPrestamosSeptiembre	.Value = 	0	;
			this.	 txtPrestamosOctubre	.Value = 	0	;
			this.	 txtPrestamosJulio	.Value = 	0	;
			this.	 txtPrestamosAgosto	.Value = 	0	;
			this.	 txtPrestamosMayo	.Value = 	0	;
			this.	 txtPrestamosJunio	.Value = 	0	;
			this.	 txtPrestamosMarzo	.Value = 	0	;
			this.	 txtPrestamosAbril	.Value = 	0	;
			this.	 txtPrestamosFebrero	.Value = 	0	;
			this.	txtPrestamosEnero	.Value = 	0	;
			this.	 txtGNCTotal	.Value = 	0	;
			this.	 txtGNCNoviembre	.Value = 	0	;
			this.	 txtGNCDiciembre	.Value = 	0	;
			this.	 txtGNCSeptiembre	.Value = 	0	;
			this.	 txtGNCOctubre	.Value = 	0	;
			this.	 txtGNCJulio	.Value = 	0	;
			this.	 txtGNCAgosto	.Value = 	0	;
			this.	 txtGNCMayo	.Value = 	0	;
			this.	 txtGNCJunio	.Value = 	0	;
			this.	 txtGNCMarzo	.Value = 	0	;
			this.	 txtGNCAbril	.Value = 	0	;
			this.	 txtGNCFebrero	.Value = 	0	;
			this.	 txtGNCEnero	.Value = 	0	;
			this.	 txtMNCTotal	.Value = 	0	;
			this.	 txtMNCNoviembre	.Value = 	0	;
			this.	 txtMNCDiciembre	.Value = 	0	;
			this.	 txtMNCSeptiembre	.Value = 	0	;
			this.	 txtMNCOctubre	.Value = 	0	;
			this.	 txtMNCJulio	.Value = 	0	;
			this.	 txtMNCAgosto	.Value = 	0	;
			this.	 txtMNCMayo	.Value = 	0	;
			this.	 txtMNCJunio	.Value = 	0	;
			this.	 txtMNCMarzo	.Value = 	0	;
			this.	 txtMNCAbril	.Value = 	0	;
			this.	 txtMNCFebrero	.Value = 	0	;
			this.	 txtMNCEnero	.Value = 	0	;
			this.	 txtSMAGastosRetirosDiciembre	.Value = 	0	;
			this.	 txtSMAGastosRetirosNoviembre	.Value = 	0	;
			this.	 txtSMAGastosRetirosOctubre	.Value = 	0	;
			this.	 txtSMAGastosRetirosSeptiembre	.Value = 	0	;
			this.	 txtSMAGastosRetirosAgosto	.Value = 	0	;
			this.	 txtSMAGastosRetirosJulio	.Value = 	0	;
			this.	 txtSMAGastosRetirosJunio	.Value = 	0	;
			this.	 txtSMAGastosRetirosMayo	.Value = 	0	;
			this.	 txtSMAGastosRetirosAbril	.Value = 	0	;
			this.	 txtSMAGastosRetirosMarzo	.Value = 	0	;
			this.	 txtSMAGastosRetirosFebrero	.Value = 	0	;
			this.	 txtTotalGastos	.Value = 	0	;
			this.	 txtTotalGastosNoviembre	.Value = 	0	;
			this.	 txtTotalGastosDiciembre	.Value = 	0	;
			this.	 txtTotalGastosSeptiembre	.Value = 	0	;
			this.	 txtTotalGastosOctubre	.Value = 	0	;
			this.	 txtTotalGastosJulio	.Value = 	0	;
			this.	 txtTotalGastosAgosto	.Value = 	0	;
			this.	 txtTotalGastosMayo	.Value = 	0	;
			this.	 txtTotalGastosJunio	.Value = 	0	;
			this.	 txtTotalGastosMarzo	.Value = 	0	;
			this.	 txtTotalGastosAbril	.Value = 	0	;
			this.	 txtTotalGastosFebrero	.Value = 	0	;
			this.	 txtTotalGastosEnero	.Value = 	0	;
			this.	 txtResultadoTotal	.Value = 	0	;
			this.	 txtResultadoNoviembre	.Value = 	0	;
			this.	txtResultadoDiciembre	.Value = 	0	;
			this.	 txtResultadoSeptiembre	.Value = 	0	;
			this.	 txtResultadoOctubre	.Value = 	0	;
			this.	 txtResultadoJulio	.Value = 	0	;
			this.	 txtResultadoAgosto	.Value = 	0	;
			this.	 txtResultadoMayo	.Value = 	0	;
			this.	 txtResultadoJunio	.Value = 	0	;
			this.	 txtResultadoMarzo	.Value = 	0	;
			this.	 txtResultadoAbril	.Value = 	0	;
			this.	 txtResultadoFebrero	.Value = 	0	;
			this.	 txtResultadoEnero	.Value = 	0	;
			this.	 txtGastosRetiroTotal	.Value = 	0	;
			this.	txtGastosRetiroNoviembre	.Value = 	0	;
			this.	 txtGastosRetiroDiciembre	.Value = 	0	;
			this.	 txtGastosRetiroSeptiembre	.Value = 	0	;
			this.	 txtGastosRetiroOctubre	.Value = 	0	;
			this.	 txtGastosRetiroJulio	.Value = 	0	;
			this.	 txtGastosRetiroAgosto	.Value = 	0	;
			this.	 txtGastosRetiroMayo	.Value = 	0	;
			this.	 txtGastosRetiroJunio	.Value = 	0	;
			this.	txtGastosRetiroMarzo	.Value = 	0	;
			this.	 txtGastosRetiroAbril	.Value = 	0	;
			this.	 txtGastosRetiroFebrero	.Value = 	0	;
			this.	 txtGastosRetiroEnero	.Value = 	0	;
			this.	 txtTotalIngresos	.Value = 	0	;
			this.	 txtTotalIngresosNoviembre	.Value = 	0	;
			this.	 txtTotalIngresosDiciembre	.Value = 	0	;
			this.	 txtTotalIngresosSeptiembre	.Value = 	0	;
			this.	 txtTotalIngresosOctubre	.Value = 	0	;
			this.	 txtTotalIngresosJulio	.Value = 	0	;
			this.	 txtTotalIngresosAgosto	.Value = 	0	;
			this.	 txtTotalIngresosMayo	.Value = 	0	;
			this.	 txtTotalIngresosJunio	.Value = 	0	;
			this.	 txtTotalIngresosMarzo	.Value = 	0	;
			this.	 txtTotalIngresosAbril	.Value = 	0	;
			this.	 txtTotalIngresosFebrero	.Value = 	0	;
			this.	 txtTotalIngresosEnero	.Value = 	0	;
			this.	 txtCobranzaTotal	.Value = 	0	;
			this.	 txtCobranzaNoviembre	.Value = 	0	;
			this.	 txtCobranzaDiciembre	.Value = 	0	;
			this.	 txtCobranzaSeptiembre	.Value = 	0	;
			this.	txtCobranzaOctubre	.Value = 	0	;
			this.	txtCobranzaJulio	.Value = 	0	;
			this.	 txtCobranzaAgosto	.Value = 	0	;
			this.	 txtCobranzaMayo	.Value = 	0	;
			this.	 txtCobranzaJunio	.Value = 	0	;
			this.	 txtCobranzaMarzo	.Value = 	0	;
			this.	 txtCobranzaAbril	.Value = 	0	;
			this.	 txtCobranzaFebrero	.Value = 	0	;
			this.	 txtCobranzaEnero	.Value = 	0	;
			this.	txtProyeccionMargenTotal	.Value = 	0	;
			this.	 txtProyeccionMargenTotalNoviembre	.Value = 	0	;
			this.	 txtProyeccionMargenTotalDiciembre	.Value = 	0	;
			this.	 txtProyeccionMargenTotalSeptiembre	.Value = 	0	;
			this.	 txtProyeccionMargenTotalOctubre	.Value = 	0	;
			this.	 txtProyeccionMargenTotalJulio	.Value = 	0	;
			this.	 txtProyeccionMargenTotalAgosto	.Value = 	0	;
			this.	 txtProyeccionMargenTotalMayo	.Value = 	0	;
			this.	txtProyeccionMargenTotalJunio	.Value = 	0	;
			this.	 txtProyeccionMargenTotalMarzo	.Value = 	0	;
			this.	 txtProyeccionMargenTotalAbril	.Value = 	0	;
			this.	 txtProyeccionMargenTotalFebrero	.Value = 	0	;
			this.	 txtProyeccionMargenTotalEnero	.Value = 	0	;
			this.	txtProyeccionVentaTotal	.Value = 	0	;
			this.	 txtProyeccionVentaTotalNoviembre	.Value = 	0	;
			this.	txtProyeccionVentaTotalDiciembre	.Value = 	0	;
			this.	 txtProyeccionVentaTotalSeptiembre	.Value = 	0	;
			this.	 txtProyeccionVentaTotalOctubre	.Value = 	0	;
			this.	 txtProyeccionVentaTotalJulio	.Value = 	0	;
			this.	 txtProyeccionVentaTotalAgosto	.Value = 	0	;
			this.	 txtProyeccionVentaTotalMayo	.Value = 	0	;
			this.	 txtProyeccionVentaTotalJunio	.Value = 	0	;
			this.	 txtProyeccionVentaTotalMarzo	.Value = 	0	;
			this.	txtProyeccionVentaTotalAbril	.Value = 	0	;
			this.	 txtProyeccionVentaTotalFebrero	.Value = 	0	;
			this.	 txtProyeccionVentaTotalEnero	.Value = 	0	;
			this.	 txtProyeccionMargenCredito	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoNoviembre	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoDiciembre	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoSeptiembre	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoOctubre	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoJulio	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoAgosto	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoMayo	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoJunio	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoMarzo	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoAbril	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoEnero	.Value = 	0	;
			this.	 txtProyeccionMargenCreditoFebrero	.Value = 	0	;
			this.	 txtProyeccionMargenTarjeta	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaNoviembre	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaDiciembre	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaSeptiembre	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaOctubre	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaJulio	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaAgosto	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaMayo	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaJunio	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaMarzo	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaAbril	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaEnero	.Value = 	0	;
			this.	 txtProyeccionMargenTarjetaFebrero	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivo	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoNoviembre	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoDiciembre	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoSeptiembre	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoOctubre	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoJulio	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoAgosto	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoMayo	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoJunio	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoMarzo	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoAbril	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoEnero	.Value = 	0	;
			this.	 txtProyeccionMargenEfectivoFebrero	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidor	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorNoviembre	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorDiciembre	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorSeptiembre	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorOctubre	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorJulio	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorAgosto	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorMayo	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorJunio	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorMarzo	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorAbril	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorEnero	.Value = 	0	;
			this.	 txtProyeccionMargenDistribuidorFebrero	.Value = 	0	;
			this.	 txtProyeccionVentaCredito	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoNoviembre	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoDiciembre	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoSeptiembre	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoOctubre	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoJulio	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoAgosto	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoMayo	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoJunio	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoMarzo	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoAbril	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoEnero	.Value = 	0	;
			this.	 txtProyeccionVentaCreditoFebrero	.Value = 	0	;
			this.	 txtProyeccionVentaTarjeta	.Value = 	0	;
			this.	 txtProyeccionVentaTarjetaNoviembre	.Value = 	0	;
			this.	txtProyeccionVentaTarjetaDiciembre	.Value = 	0	;
			this.	 txtProyeccionVentaTarjetaSeptiembre	.Value = 	0	;
			this.	txtProyeccionVentaTarjetaOctubre	.Value = 	0	;
			this.	 txtProyeccionVentaTarjetaJulio	.Value = 	0	;
			this.	 txtProyeccionVentaTarjetaAgosto	.Value = 	0	;
			this.	 txtProyeccionVentaTarjetaMayo	.Value = 	0	;
			this.	 txtProyeccionVentaTarjetaJunio	.Value = 	0	;
			this.	 txtProyeccionVentaTarjetaMarzo	.Value = 	0	;
			this.	 txtProyeccionVentaTarjetaAbril	.Value = 	0	;
			this.	 txtProyeccionVentaTarjetaEnero	.Value = 	0	;
			this.	 txtProyeccionVentaTarjetaFebrero	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivo	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivoNoviembre	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivoDiciembre	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivoSeptiembre	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivoOctubre	.Value = 	0	;
			this.	txtProyeccionVentaEfectivoJulio	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivoAgosto	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivoMayo	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivoJunio	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivoMarzo	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivoAbril	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivoEnero	.Value = 	0	;
			this.	 txtProyeccionVentaEfectivoFebrero	.Value = 	0	;
			this.	 txtProyeccionVentaDistribuidor	.Value = 	0	;
			this.	 txtIngresosTotal	.Value = 	0	;
			this.	 txtProyeccionVentaDistribuidorNoviembre	.Value = 	0	;
			this.	 txtProyeccionVentaDistribuidorDiciembre	.Value = 	0	;
			this.	 txtProyeccionDiciembreVenta	.Value = 	0	;
			this.	 txtProyeccionNoviembreVenta	.Value = 	0	;
			this.	 txtProyeccionVentaDistribuidorSeptiembre	.Value = 	0	;
			this.	 txtProyeccionVentaDistribuidorOctubre	.Value = 	0	;
			this.	txtProyeccionOctubreVenta	.Value = 	0	;
			this.	txtProyeccionSeptiembreVenta	.Value = 	0	;
			this.	 txtProyeccionVentaDistribuidorJulio	.Value = 	0	;
			this.	 txtProyeccionVentaDistribuidorAgosto	.Value = 	0	;
			this.	 txtProyeccionAgostoVenta	.Value = 	0	;
			this.	 txtProyeccionJulioVenta	.Value = 	0	;
			this.	 txtProyeccionVentaDistribuidorMayo	.Value = 	0	;
			this.	 txtProyeccionVentaDistribuidorJunio	.Value = 	0	;
			this.	 txtProyeccionJunioVenta	.Value = 	0	;
			this.	 txtProyeccionMayoVenta	.Value = 	0	;
			this.	 txtProyeccionVentaDistribuidorMarzo	.Value = 	0	;
			this.	 txtProyeccionVentaDistribuidorAbril	.Value = 	0	;
			this.	 txtProyeccionAbrilVenta	.Value = 	0	;
			this.txtProyeccionMarzoVenta	.Value = 	0	;
			this.txtProyeccionVentaDistribuidorEnero	.Value = 	0	;
			this.txtProyeccionVentaDistribuidorFebrero	.Value = 	0	;
			this.txtProyeccionFebreroVenta	.Value = 	0	;
			this.txtProyeccionEneroVenta	.Value = 	0	;

		}

		private void AbreResumenRetiros(int iMes, int iAño)
		{
//			frmProyeccionResumenRetiros PRR = new frmProyeccionResumenRetiros(iMes, iAño);
//			PRR.ShowDialog();
		}

		private void txtGNCEnero_DoubleClick(object sender, System.EventArgs e)
		{
		 AbreDocumentos(1, iPeriodo, 1);
		}

		private void txtGNCFebrero_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(2, iPeriodo, 1);
		}

		private void txtGNCMarzo_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(3, iPeriodo, 1);
		}

		private void txtGNCAbril_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(4, iPeriodo, 1);
		}

		private void txtGNCMayo_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(5, iPeriodo, 1);
		}

		private void txtGNCJunio_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(6, iPeriodo, 1);
		}

		private void txtGNCJulio_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(7, iPeriodo, 1);
		}

		private void txtGNCAgosto_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(8, iPeriodo, 1);
		}

		private void txtGNCSeptiembre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(9, iPeriodo, 1);
		}

		private void txtGNCOctubre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(10, iPeriodo, 1);
		}

		private void txtGNCNoviembre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(11, iPeriodo, 1);
		}

		private void txtGNCDiciembre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(12, iPeriodo, 1);
		}

		private void txtGNCTotal_DoubleClick(object sender, System.EventArgs e)
		{
		
		}

		private void txtMNCEnero_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(1, iPeriodo, 2);
		}

		private void txtMNCFebrero_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(2, iPeriodo, 2);
		}

		private void txtMNCMarzo_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(3, iPeriodo, 2);
		}

		private void txtMNCAbril_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(4, iPeriodo, 2);
		}

		private void txtMNCMayo_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(5, iPeriodo, 2);
		}

		private void txtMNCJunio_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(6, iPeriodo, 2);
		}

		private void txtMNCJulio_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(7, iPeriodo, 2);
		}

		private void txtMNCAgosto_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(8, iPeriodo, 2);
		}

		private void txtMNCSeptiembre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(9, iPeriodo, 2);
		}

		private void txtMNCOctubre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(10, iPeriodo, 2);
		}

		private void txtMNCNoviembre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(11, iPeriodo, 2);
		}

		private void txtMNCDiciembre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreDocumentos(12, iPeriodo, 2);
		}

		private void txtMNCTotal_DoubleClick(object sender, System.EventArgs e)
		{
		
		}

		private void txtGastosRetiroEnero_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(1, iPeriodo);
		}

		private void txtGastosRetiroFebrero_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(2, iPeriodo);
		}

		private void txtGastosRetiroMarzo_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(3, iPeriodo);
		}

		private void txtGastosRetiroAbril_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(4, iPeriodo);
		}

		private void txtGastosRetiroMayo_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(5, iPeriodo);
		}

		private void txtGastosRetiroJunio_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(6, iPeriodo);
		}

		private void txtGastosRetiroJulio_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(7, iPeriodo);
		}

		private void txtGastosRetiroAgosto_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(8, iPeriodo);
		}

		private void txtGastosRetiroSeptiembre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(9, iPeriodo);
		}

		private void txtGastosRetiroOctubre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(10, iPeriodo);
		}

		private void txtGastosRetiroNoviembre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(11, iPeriodo);
		}

		private void txtGastosRetiroDiciembre_DoubleClick(object sender, System.EventArgs e)
		{
		AbreResumenRetiros(12, iPeriodo);
		}

		private void txtGastosRetiroTotal_DoubleClick(object sender, System.EventArgs e)
		{
		
		}

		private void txtGNCJulio_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtGNCAgosto_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor3_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor4_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor5_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor6_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			vaciar ();
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una Bodega")) return;
			 iPeriodo = (int) this.txtAño.Value;
		  this.ConsultaGeneral();
      

		}


		private void txtAño_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void uGridGastos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void optContrato_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click (sender, e);
		}

		

	}
}
