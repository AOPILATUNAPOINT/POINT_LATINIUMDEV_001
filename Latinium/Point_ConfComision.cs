using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Point_ConfComision.
	/// </summary>
	public class Point_ConfComision : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaDesde;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl pgEntrega;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcTransaccion;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private C1.Data.C1DataSet cdsComision;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdComisiones;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private System.Windows.Forms.ComboBox cmbEstado;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdSueldos;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEvalua5;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAplicae5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEvalua4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAplicae4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEvalua3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAplicae3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEvalua2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAplicae2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEvalua1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAplicae1;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbLocal;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdMetas;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEjecutivo;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdRolEjecutivo;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtComision;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRolVen;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoVenta;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTPCalificacion;
		private Infragistics.Win.Misc.UltraButton btnTransforma;
		private DevExpress.Utils.ToolTipController toolTipController1;
		private System.Windows.Forms.ToolTip toolTip;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.ComponentModel.IContainer components;

		public Point_ConfComision()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNwBodg");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cod_UB");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Orden", 0);
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Com_Tipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRegistro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FecDesde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FecHasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Status");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo", -1, "cmbTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoLU", -1, "cmbLocal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cant_Vend");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PctContado");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotaContado");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PctCredito");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotaCredito");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotaZense");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolVendedores", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RolVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CODTRIM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbEjecutivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaInicia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaFin");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiendas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolVentas", -1, "cmbRolVen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cant_Dias");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CODTRIM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaInicia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaFin");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiendas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cant_Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Com_TipoCalificacion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Com_TipoVenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoVenta", -1, "cmbTipoVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalificacion", -1, "cmbTPCalificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VMinimo");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VMaximo");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VGrteVentas");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VJefeVentas");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VSupervisor");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VJefeTiendaA");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VJefeTiendaB");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VJefeTiendaC");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VVendedor");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPorcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VMinimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VMaximo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VGrteVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VJefeVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VSupervisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VJefeTiendaA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VJefeTiendaB");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VJefeTiendaC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VVendedor");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoUsuario");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoUsuario");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoUsuario");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoUsuario");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolVendedores", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RolVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoUsuario");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab(true);
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Point_ConfComision));
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbLocal = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsComision = new C1.Data.C1DataSet();
			this.grdMetas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbRolVen = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCargo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbEjecutivo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.grdRolEjecutivo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbTPCalificacion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbTipoVenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.grdComisiones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.pgEntrega = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbEvalua5 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label8 = new System.Windows.Forms.Label();
			this.chkAplicae5 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbEvalua4 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.chkAplicae4 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbEvalua3 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.chkAplicae3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbEvalua2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.chkAplicae2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.grdSueldos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbEvalua1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label38 = new System.Windows.Forms.Label();
			this.chkAplicae1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFechaHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.utcTransaccion = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.cmbEstado = new System.Windows.Forms.ComboBox();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtComision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnTransforma = new Infragistics.Win.Misc.UltraButton();
			this.toolTipController1 = new DevExpress.Utils.ToolTipController(this.components);
			this.toolTip = new System.Windows.Forms.ToolTip(this.components);
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdMetas)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbRolVen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEjecutivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRolEjecutivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTPCalificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdComisiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.pgEntrega.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbEvalua5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEvalua4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEvalua3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEvalua2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdSueldos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEvalua1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcTransaccion)).BeginInit();
			this.utcTransaccion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.cmbLocal);
			this.ultraTabPageControl1.Controls.Add(this.cmbTipo);
			this.ultraTabPageControl1.Controls.Add(this.grdMetas);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1214, 396);
			this.toolTipController1.SetSuperTip(this.ultraTabPageControl1, null);
			// 
			// cmbLocal
			// 
			this.cmbLocal.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 300;
			ultraGridColumn3.Header.Caption = "Codigo";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 70;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbLocal.DisplayMember = "Nombre";
			this.cmbLocal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLocal.Location = new System.Drawing.Point(192, 72);
			this.cmbLocal.Name = "cmbLocal";
			this.cmbLocal.Size = new System.Drawing.Size(176, 86);
			this.toolTipController1.SetSuperTip(this.cmbLocal, null);
			this.cmbLocal.TabIndex = 56;
			this.cmbLocal.ValueMember = "Cod_UB";
			this.cmbLocal.Visible = false;
			// 
			// cmbTipo
			// 
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataMember = "Com_Tipo";
			this.cmbTipo.DataSource = this.cdsComision;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipo.DisplayMember = "Codigo";
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(32, 72);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(152, 86);
			this.toolTipController1.SetSuperTip(this.cmbTipo, null);
			this.cmbTipo.TabIndex = 55;
			this.cmbTipo.ValueMember = "idTipo";
			this.cmbTipo.Visible = false;
			// 
			// cdsComision
			// 
			this.cdsComision.BindingContextCtrl = this;
			this.cdsComision.DataLibrary = "LibFacturacion";
			this.cdsComision.DataLibraryUrl = "";
			this.cdsComision.DataSetDef = "dsComisiones";
			this.cdsComision.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsComision.Name = "cdsComision";
			this.cdsComision.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsComision.SchemaDef = null;
			this.cdsComision.BeforeFill += new C1.Data.FillEventHandler(this.cdsComision_BeforeFill);
			// 
			// grdMetas
			// 
			this.grdMetas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdMetas.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdMetas.DisplayLayout.Appearance = appearance1;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Width = 100;
			ultraGridColumn15.Header.Caption = "Local / Usuario";
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Width = 400;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance2;
			ultraGridColumn16.Header.Caption = "# Vendres";
			ultraGridColumn16.Header.VisiblePosition = 9;
			ultraGridColumn16.Width = 50;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance3;
			ultraGridColumn17.Header.VisiblePosition = 10;
			ultraGridColumn17.Width = 70;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance4;
			ultraGridColumn18.Header.Caption = "% Contado";
			ultraGridColumn18.Header.VisiblePosition = 11;
			ultraGridColumn18.Width = 70;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance5;
			ultraGridColumn19.Header.Caption = "Val Contado";
			ultraGridColumn19.Header.VisiblePosition = 12;
			ultraGridColumn19.Width = 70;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance6;
			ultraGridColumn20.Header.Caption = "% Crédito";
			ultraGridColumn20.Header.VisiblePosition = 13;
			ultraGridColumn20.Width = 70;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance7;
			ultraGridColumn21.Header.Caption = "Val Crédito";
			ultraGridColumn21.Header.VisiblePosition = 14;
			ultraGridColumn21.Width = 70;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance8;
			ultraGridColumn22.Header.Caption = "Zense";
			ultraGridColumn22.Header.VisiblePosition = 15;
			ultraGridColumn22.Width = 70;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn23.CellAppearance = appearance9;
			ultraGridColumn23.Header.Caption = "# Días";
			ultraGridColumn23.Header.VisiblePosition = 16;
			ultraGridColumn23.Width = 40;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn23});
			this.grdMetas.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.grdMetas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdMetas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.grdMetas.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 10F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdMetas.DisplayLayout.Override.HeaderAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdMetas.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdMetas.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.grdMetas.Location = new System.Drawing.Point(7, 5);
			this.grdMetas.Name = "grdMetas";
			this.grdMetas.Size = new System.Drawing.Size(1200, 385);
			this.toolTipController1.SetSuperTip(this.grdMetas, null);
			this.grdMetas.TabIndex = 58;
			this.grdMetas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.grdMetas_KeyPress);
			this.grdMetas.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.grdMetas_BeforeRowsDeleted);
			this.grdMetas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdMetas_KeyDown);
			this.grdMetas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdMetas_AfterCellUpdate_1);
			this.grdMetas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdMetas_InitializeLayout);
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.cmbRolVen);
			this.ultraTabPageControl2.Controls.Add(this.cmbCargo);
			this.ultraTabPageControl2.Controls.Add(this.cmbEjecutivo);
			this.ultraTabPageControl2.Controls.Add(this.grdRolEjecutivo);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(1214, 396);
			this.toolTipController1.SetSuperTip(this.ultraTabPageControl2, null);
			// 
			// cmbRolVen
			// 
			this.cmbRolVen.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRolVen.DataMember = "RolVendedores";
			this.cmbRolVen.DataSource = this.cdsComision;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 200;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbRolVen.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbRolVen.DisplayMember = "RolVentas";
			this.cmbRolVen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRolVen.Location = new System.Drawing.Point(456, 96);
			this.cmbRolVen.Name = "cmbRolVen";
			this.cmbRolVen.Size = new System.Drawing.Size(176, 86);
			this.toolTipController1.SetSuperTip(this.cmbRolVen, null);
			this.cmbRolVen.TabIndex = 60;
			this.cmbRolVen.ValueMember = "idRolVentas";
			this.cmbRolVen.Visible = false;
			// 
			// cmbCargo
			// 
			this.cmbCargo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 150;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCargo.DisplayMember = "Nombre";
			this.cmbCargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargo.Location = new System.Drawing.Point(256, 96);
			this.cmbCargo.Name = "cmbCargo";
			this.cmbCargo.Size = new System.Drawing.Size(176, 86);
			this.toolTipController1.SetSuperTip(this.cmbCargo, null);
			this.cmbCargo.TabIndex = 59;
			this.cmbCargo.ValueMember = "idCargo";
			this.cmbCargo.Visible = false;
			// 
			// cmbEjecutivo
			// 
			this.cmbEjecutivo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 400;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.cmbEjecutivo.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbEjecutivo.DisplayMember = "Nombre";
			this.cmbEjecutivo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEjecutivo.Location = new System.Drawing.Point(56, 96);
			this.cmbEjecutivo.Name = "cmbEjecutivo";
			this.cmbEjecutivo.Size = new System.Drawing.Size(176, 86);
			this.toolTipController1.SetSuperTip(this.cmbEjecutivo, null);
			this.cmbEjecutivo.TabIndex = 57;
			this.cmbEjecutivo.ValueMember = "idPersonal";
			this.cmbEjecutivo.Visible = false;
			// 
			// grdRolEjecutivo
			// 
			this.grdRolEjecutivo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdRolEjecutivo.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdRolEjecutivo.DataSource = this.ultraDataSource1;
			appearance14.BackColor = System.Drawing.Color.White;
			this.grdRolEjecutivo.DisplayLayout.Appearance = appearance14;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 2;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.Caption = "Ejecutivo";
			ultraGridColumn34.Header.VisiblePosition = 3;
			ultraGridColumn34.Width = 380;
			ultraGridColumn35.Header.Caption = "Cargo";
			ultraGridColumn35.Header.VisiblePosition = 4;
			ultraGridColumn36.Header.Caption = "Inicia";
			ultraGridColumn36.Header.VisiblePosition = 5;
			ultraGridColumn36.Width = 80;
			ultraGridColumn37.Header.Caption = "Finaliza";
			ultraGridColumn37.Header.VisiblePosition = 6;
			ultraGridColumn37.Width = 80;
			ultraGridColumn38.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn38.Header.VisiblePosition = 7;
			ultraGridColumn38.Width = 200;
			ultraGridColumn39.Header.Caption = "Rol Ventas";
			ultraGridColumn39.Header.VisiblePosition = 8;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn40.CellAppearance = appearance15;
			ultraGridColumn40.Header.Caption = "Dias";
			ultraGridColumn40.Header.VisiblePosition = 9;
			ultraGridColumn40.Width = 40;
			ultraGridColumn41.Header.VisiblePosition = 10;
			ultraGridColumn41.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn42.Header.VisiblePosition = 11;
			ultraGridBand7.Columns.AddRange(new object[] {
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
			this.grdRolEjecutivo.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.grdRolEjecutivo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.grdRolEjecutivo.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 10F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdRolEjecutivo.DisplayLayout.Override.HeaderAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRolEjecutivo.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRolEjecutivo.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.grdRolEjecutivo.Location = new System.Drawing.Point(7, 5);
			this.grdRolEjecutivo.Name = "grdRolEjecutivo";
			this.grdRolEjecutivo.Size = new System.Drawing.Size(1200, 385);
			this.toolTipController1.SetSuperTip(this.grdRolEjecutivo, null);
			this.grdRolEjecutivo.TabIndex = 58;
			this.grdRolEjecutivo.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.grdRolEjecutivo_BeforeRowsDeleted);
			this.grdRolEjecutivo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdRolEjecutivo_KeyDown_1);
			this.grdRolEjecutivo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdRolEjecutivo_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn11.DataType = typeof(int);
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
																																 ultraDataColumn12});
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.cmbTPCalificacion);
			this.ultraTabPageControl3.Controls.Add(this.cmbTipoVenta);
			this.ultraTabPageControl3.Controls.Add(this.grdComisiones);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(1, 20);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(1214, 396);
			this.toolTipController1.SetSuperTip(this.ultraTabPageControl3, null);
			// 
			// cmbTPCalificacion
			// 
			this.cmbTPCalificacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTPCalificacion.DataMember = "Com_TipoCalificacion";
			this.cmbTPCalificacion.DataSource = this.cdsComision;
			ultraGridBand8.AddButtonCaption = "RolVendedores";
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Width = 150;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn43,
																										 ultraGridColumn44});
			this.cmbTPCalificacion.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbTPCalificacion.DisplayMember = "Nombre";
			this.cmbTPCalificacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTPCalificacion.Location = new System.Drawing.Point(232, 104);
			this.cmbTPCalificacion.Name = "cmbTPCalificacion";
			this.cmbTPCalificacion.Size = new System.Drawing.Size(176, 86);
			this.toolTipController1.SetSuperTip(this.cmbTPCalificacion, null);
			this.cmbTPCalificacion.TabIndex = 62;
			this.cmbTPCalificacion.ValueMember = "idTipoCalificacion";
			this.cmbTPCalificacion.Visible = false;
			// 
			// cmbTipoVenta
			// 
			this.cmbTipoVenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoVenta.DataMember = "Com_TipoVenta";
			this.cmbTipoVenta.DataSource = this.cdsComision;
			ultraGridBand9.AddButtonCaption = "RolVendedores";
			ultraGridColumn45.Header.VisiblePosition = 0;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 1;
			ultraGridColumn46.Width = 150;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn45,
																										 ultraGridColumn46});
			this.cmbTipoVenta.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbTipoVenta.DisplayMember = "Nombre";
			this.cmbTipoVenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoVenta.Location = new System.Drawing.Point(40, 104);
			this.cmbTipoVenta.Name = "cmbTipoVenta";
			this.cmbTipoVenta.Size = new System.Drawing.Size(176, 86);
			this.toolTipController1.SetSuperTip(this.cmbTipoVenta, null);
			this.cmbTipoVenta.TabIndex = 61;
			this.cmbTipoVenta.ValueMember = "idTipoVenta";
			this.cmbTipoVenta.Visible = false;
			// 
			// grdComisiones
			// 
			this.grdComisiones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdComisiones.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdComisiones.DataSource = this.ultraDataSource2;
			appearance20.BackColor = System.Drawing.Color.White;
			this.grdComisiones.DisplayLayout.Appearance = appearance20;
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 1;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.Caption = "Venta";
			ultraGridColumn49.Header.VisiblePosition = 2;
			ultraGridColumn49.Width = 198;
			ultraGridColumn50.Header.Caption = "Calificación";
			ultraGridColumn50.Header.VisiblePosition = 3;
			ultraGridColumn50.Width = 149;
			ultraGridColumn51.Header.VisiblePosition = 4;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 250;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn52.CellAppearance = appearance21;
			ultraGridColumn52.Header.Caption = "Mínimo";
			ultraGridColumn52.Header.VisiblePosition = 5;
			ultraGridColumn52.Width = 60;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn53.CellAppearance = appearance22;
			ultraGridColumn53.Header.Caption = "Máximo";
			ultraGridColumn53.Header.VisiblePosition = 6;
			ultraGridColumn53.Width = 60;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn54.CellAppearance = appearance23;
			ultraGridColumn54.Header.Caption = "Grte. Ventas";
			ultraGridColumn54.Header.VisiblePosition = 7;
			ultraGridColumn54.Width = 80;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn55.CellAppearance = appearance24;
			ultraGridColumn55.Header.Caption = "Jefe Ventas";
			ultraGridColumn55.Header.VisiblePosition = 8;
			ultraGridColumn55.Width = 80;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn56.CellAppearance = appearance25;
			ultraGridColumn56.Header.Caption = "Supervisor";
			ultraGridColumn56.Header.VisiblePosition = 9;
			ultraGridColumn56.Width = 80;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn57.CellAppearance = appearance26;
			ultraGridColumn57.Header.Caption = "Jefe Tienda A";
			ultraGridColumn57.Header.VisiblePosition = 10;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn58.CellAppearance = appearance27;
			ultraGridColumn58.Header.Caption = "Jefe Tienda B";
			ultraGridColumn58.Header.VisiblePosition = 11;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn59.CellAppearance = appearance28;
			ultraGridColumn59.Header.Caption = "Jefe Tienda C";
			ultraGridColumn59.Header.VisiblePosition = 12;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn60.CellAppearance = appearance29;
			ultraGridColumn60.Header.Caption = "Vendedor";
			ultraGridColumn60.Header.VisiblePosition = 13;
			ultraGridColumn60.Width = 80;
			ultraGridBand10.Columns.AddRange(new object[] {
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
																											ultraGridColumn60});
			this.grdComisiones.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.grdComisiones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance30.BackColor = System.Drawing.Color.Transparent;
			this.grdComisiones.DisplayLayout.Override.CardAreaAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance31.FontData.BoldAsString = "True";
			appearance31.FontData.Name = "Arial";
			appearance31.FontData.SizeInPoints = 10F;
			appearance31.ForeColor = System.Drawing.Color.White;
			appearance31.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdComisiones.DisplayLayout.Override.HeaderAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdComisiones.DisplayLayout.Override.RowSelectorAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdComisiones.DisplayLayout.Override.SelectedRowAppearance = appearance33;
			this.grdComisiones.Location = new System.Drawing.Point(8, 8);
			this.grdComisiones.Name = "grdComisiones";
			this.grdComisiones.Size = new System.Drawing.Size(1200, 385);
			this.toolTipController1.SetSuperTip(this.grdComisiones, null);
			this.grdComisiones.TabIndex = 10;
			this.grdComisiones.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.grdComisiones_InitializeRow);
			this.grdComisiones.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.grdComisiones_BeforeRowsDeleted);
			this.grdComisiones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdComisiones_KeyDown);
			this.grdComisiones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdComisiones_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14,
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
																																 ultraDataColumn26});
			// 
			// pgEntrega
			// 
			this.pgEntrega.Controls.Add(this.cmbEvalua5);
			this.pgEntrega.Controls.Add(this.label8);
			this.pgEntrega.Controls.Add(this.chkAplicae5);
			this.pgEntrega.Controls.Add(this.cmbEvalua4);
			this.pgEntrega.Controls.Add(this.label7);
			this.pgEntrega.Controls.Add(this.chkAplicae4);
			this.pgEntrega.Controls.Add(this.cmbEvalua3);
			this.pgEntrega.Controls.Add(this.label6);
			this.pgEntrega.Controls.Add(this.chkAplicae3);
			this.pgEntrega.Controls.Add(this.cmbEvalua2);
			this.pgEntrega.Controls.Add(this.label5);
			this.pgEntrega.Controls.Add(this.chkAplicae2);
			this.pgEntrega.Controls.Add(this.grdSueldos);
			this.pgEntrega.Controls.Add(this.cmbEvalua1);
			this.pgEntrega.Controls.Add(this.label38);
			this.pgEntrega.Controls.Add(this.chkAplicae1);
			this.pgEntrega.Location = new System.Drawing.Point(-10000, -10000);
			this.pgEntrega.Name = "pgEntrega";
			this.pgEntrega.Size = new System.Drawing.Size(1214, 396);
			this.toolTipController1.SetSuperTip(this.pgEntrega, null);
			// 
			// cmbEvalua5
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEvalua5.Appearance = appearance34;
			this.cmbEvalua5.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEvalua5.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEvalua5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsComision, "Com_CtaComision.Evaluador_5"));
			ultraGridColumn61.Header.VisiblePosition = 0;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 1;
			ultraGridColumn62.Width = 300;
			ultraGridColumn63.Header.VisiblePosition = 2;
			ultraGridColumn63.Width = 70;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn61,
																											ultraGridColumn62,
																											ultraGridColumn63});
			this.cmbEvalua5.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbEvalua5.DisplayMember = "Nombre";
			this.cmbEvalua5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEvalua5.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEvalua5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEvalua5.Location = new System.Drawing.Point(104, 112);
			this.cmbEvalua5.Name = "cmbEvalua5";
			this.cmbEvalua5.Size = new System.Drawing.Size(256, 21);
			this.toolTipController1.SetSuperTip(this.cmbEvalua5, null);
			this.cmbEvalua5.TabIndex = 89;
			this.cmbEvalua5.ValueMember = "idPersonal";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(32, 114);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(68, 16);
			this.toolTipController1.SetSuperTip(this.label8, null);
			this.label8.TabIndex = 88;
			this.label8.Text = "Evaluador 5:";
			// 
			// chkAplicae5
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAplicae5.Appearance = appearance35;
			this.chkAplicae5.BackColor = System.Drawing.Color.Transparent;
			this.chkAplicae5.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsComision, "Com_CtaComision.Aplica_5"));
			this.chkAplicae5.Location = new System.Drawing.Point(376, 111);
			this.chkAplicae5.Name = "chkAplicae5";
			this.chkAplicae5.Size = new System.Drawing.Size(120, 22);
			this.toolTipController1.SetSuperTip(this.chkAplicae5, null);
			this.chkAplicae5.TabIndex = 87;
			this.chkAplicae5.Text = "Aplica Evaluación";
			// 
			// cmbEvalua4
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEvalua4.Appearance = appearance36;
			this.cmbEvalua4.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEvalua4.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEvalua4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsComision, "Com_CtaComision.Evaluador_4"));
			ultraGridColumn64.Header.VisiblePosition = 0;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 1;
			ultraGridColumn65.Width = 300;
			ultraGridColumn66.Header.VisiblePosition = 2;
			ultraGridColumn66.Width = 70;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn64,
																											ultraGridColumn65,
																											ultraGridColumn66});
			this.cmbEvalua4.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbEvalua4.DisplayMember = "Nombre";
			this.cmbEvalua4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEvalua4.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEvalua4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEvalua4.Location = new System.Drawing.Point(104, 88);
			this.cmbEvalua4.Name = "cmbEvalua4";
			this.cmbEvalua4.Size = new System.Drawing.Size(256, 21);
			this.toolTipController1.SetSuperTip(this.cmbEvalua4, null);
			this.cmbEvalua4.TabIndex = 86;
			this.cmbEvalua4.ValueMember = "idPersonal";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(32, 90);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(68, 16);
			this.toolTipController1.SetSuperTip(this.label7, null);
			this.label7.TabIndex = 85;
			this.label7.Text = "Evaluador 4:";
			// 
			// chkAplicae4
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAplicae4.Appearance = appearance37;
			this.chkAplicae4.BackColor = System.Drawing.Color.Transparent;
			this.chkAplicae4.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsComision, "Com_CtaComision.Aplica_4"));
			this.chkAplicae4.Location = new System.Drawing.Point(376, 87);
			this.chkAplicae4.Name = "chkAplicae4";
			this.chkAplicae4.Size = new System.Drawing.Size(120, 22);
			this.toolTipController1.SetSuperTip(this.chkAplicae4, null);
			this.chkAplicae4.TabIndex = 84;
			this.chkAplicae4.Text = "Aplica Evaluación";
			// 
			// cmbEvalua3
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEvalua3.Appearance = appearance38;
			this.cmbEvalua3.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEvalua3.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEvalua3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsComision, "Com_CtaComision.Evaluador_3"));
			ultraGridColumn67.Header.VisiblePosition = 0;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 1;
			ultraGridColumn68.Width = 300;
			ultraGridColumn69.Header.VisiblePosition = 2;
			ultraGridColumn69.Width = 70;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn67,
																											ultraGridColumn68,
																											ultraGridColumn69});
			this.cmbEvalua3.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbEvalua3.DisplayMember = "Nombre";
			this.cmbEvalua3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEvalua3.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEvalua3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEvalua3.Location = new System.Drawing.Point(104, 64);
			this.cmbEvalua3.Name = "cmbEvalua3";
			this.cmbEvalua3.Size = new System.Drawing.Size(256, 21);
			this.toolTipController1.SetSuperTip(this.cmbEvalua3, null);
			this.cmbEvalua3.TabIndex = 83;
			this.cmbEvalua3.ValueMember = "idPersonal";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(32, 66);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(68, 16);
			this.toolTipController1.SetSuperTip(this.label6, null);
			this.label6.TabIndex = 82;
			this.label6.Text = "Evaluador 3:";
			// 
			// chkAplicae3
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAplicae3.Appearance = appearance39;
			this.chkAplicae3.BackColor = System.Drawing.Color.Transparent;
			this.chkAplicae3.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsComision, "Com_CtaComision.Aplica_3"));
			this.chkAplicae3.Location = new System.Drawing.Point(376, 63);
			this.chkAplicae3.Name = "chkAplicae3";
			this.chkAplicae3.Size = new System.Drawing.Size(120, 22);
			this.toolTipController1.SetSuperTip(this.chkAplicae3, null);
			this.chkAplicae3.TabIndex = 81;
			this.chkAplicae3.Text = "Aplica Evaluación";
			// 
			// cmbEvalua2
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEvalua2.Appearance = appearance40;
			this.cmbEvalua2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEvalua2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEvalua2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsComision, "Com_CtaComision.Evaluador_2"));
			ultraGridColumn70.Header.VisiblePosition = 0;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 1;
			ultraGridColumn71.Width = 300;
			ultraGridColumn72.Header.VisiblePosition = 2;
			ultraGridColumn72.Width = 70;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn70,
																											ultraGridColumn71,
																											ultraGridColumn72});
			this.cmbEvalua2.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbEvalua2.DisplayMember = "Nombre";
			this.cmbEvalua2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEvalua2.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEvalua2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEvalua2.Location = new System.Drawing.Point(104, 40);
			this.cmbEvalua2.Name = "cmbEvalua2";
			this.cmbEvalua2.Size = new System.Drawing.Size(256, 21);
			this.toolTipController1.SetSuperTip(this.cmbEvalua2, null);
			this.cmbEvalua2.TabIndex = 80;
			this.cmbEvalua2.ValueMember = "idPersonal";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(32, 42);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(68, 16);
			this.toolTipController1.SetSuperTip(this.label5, null);
			this.label5.TabIndex = 79;
			this.label5.Text = "Evaluador 2:";
			// 
			// chkAplicae2
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAplicae2.Appearance = appearance41;
			this.chkAplicae2.BackColor = System.Drawing.Color.Transparent;
			this.chkAplicae2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsComision, "Com_CtaComision.Aplica_2"));
			this.chkAplicae2.Location = new System.Drawing.Point(376, 39);
			this.chkAplicae2.Name = "chkAplicae2";
			this.chkAplicae2.Size = new System.Drawing.Size(120, 22);
			this.toolTipController1.SetSuperTip(this.chkAplicae2, null);
			this.chkAplicae2.TabIndex = 78;
			this.chkAplicae2.Text = "Aplica Evaluación";
			// 
			// grdSueldos
			// 
			this.grdSueldos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdSueldos.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdSueldos.DataMember = "RolVendedores";
			this.grdSueldos.DataSource = this.cdsComision;
			appearance42.BackColor = System.Drawing.Color.White;
			this.grdSueldos.DisplayLayout.Appearance = appearance42;
			ultraGridColumn73.Header.VisiblePosition = 0;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.Caption = "Cargo";
			ultraGridColumn74.Header.VisiblePosition = 1;
			ultraGridColumn74.Width = 400;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance43;
			ultraGridColumn75.Header.VisiblePosition = 2;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn73,
																											ultraGridColumn74,
																											ultraGridColumn75});
			this.grdSueldos.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.grdSueldos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance44.BackColor = System.Drawing.Color.Transparent;
			this.grdSueldos.DisplayLayout.Override.CardAreaAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance45.FontData.BoldAsString = "True";
			appearance45.FontData.Name = "Arial";
			appearance45.FontData.SizeInPoints = 10F;
			appearance45.ForeColor = System.Drawing.Color.White;
			appearance45.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdSueldos.DisplayLayout.Override.HeaderAppearance = appearance45;
			appearance46.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance46.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSueldos.DisplayLayout.Override.RowSelectorAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSueldos.DisplayLayout.Override.SelectedRowAppearance = appearance47;
			this.grdSueldos.Location = new System.Drawing.Point(544, 5);
			this.grdSueldos.Name = "grdSueldos";
			this.grdSueldos.Size = new System.Drawing.Size(656, 385);
			this.toolTipController1.SetSuperTip(this.grdSueldos, null);
			this.grdSueldos.TabIndex = 77;
			// 
			// cmbEvalua1
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEvalua1.Appearance = appearance48;
			this.cmbEvalua1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEvalua1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEvalua1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsComision, "Com_CtaComision.Evaluador_1"));
			ultraGridColumn76.Header.VisiblePosition = 0;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 1;
			ultraGridColumn77.Width = 300;
			ultraGridColumn78.Header.VisiblePosition = 2;
			ultraGridColumn78.Width = 70;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn76,
																											ultraGridColumn77,
																											ultraGridColumn78});
			this.cmbEvalua1.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbEvalua1.DisplayMember = "Nombre";
			this.cmbEvalua1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEvalua1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEvalua1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEvalua1.Location = new System.Drawing.Point(104, 17);
			this.cmbEvalua1.Name = "cmbEvalua1";
			this.cmbEvalua1.Size = new System.Drawing.Size(256, 21);
			this.toolTipController1.SetSuperTip(this.cmbEvalua1, null);
			this.cmbEvalua1.TabIndex = 76;
			this.cmbEvalua1.ValueMember = "idPersonal";
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.Color.Transparent;
			this.label38.Location = new System.Drawing.Point(32, 19);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(68, 16);
			this.toolTipController1.SetSuperTip(this.label38, null);
			this.label38.TabIndex = 72;
			this.label38.Text = "Evaluador 1:";
			// 
			// chkAplicae1
			// 
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAplicae1.Appearance = appearance49;
			this.chkAplicae1.BackColor = System.Drawing.Color.Transparent;
			this.chkAplicae1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsComision, "Com_CtaComision.Aplica_1"));
			this.chkAplicae1.Location = new System.Drawing.Point(376, 16);
			this.chkAplicae1.Name = "chkAplicae1";
			this.chkAplicae1.Size = new System.Drawing.Size(120, 22);
			this.toolTipController1.SetSuperTip(this.chkAplicae1, null);
			this.chkAplicae1.TabIndex = 52;
			this.chkAplicae1.Text = "Aplica Evaluación";
			// 
			// dtFechaHasta
			// 
			this.dtFechaHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaHasta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsComision, "Com_CtaComision.FecHasta"));
			dateButton1.Caption = "Today";
			this.dtFechaHasta.DateButtons.Add(dateButton1);
			this.dtFechaHasta.Format = "dd/MMM/yyyy";
			this.dtFechaHasta.Location = new System.Drawing.Point(624, 32);
			this.dtFechaHasta.Name = "dtFechaHasta";
			this.dtFechaHasta.NonAutoSizeHeight = 23;
			this.dtFechaHasta.Size = new System.Drawing.Size(112, 21);
			this.dtFechaHasta.SpinButtonsVisible = true;
			this.toolTipController1.SetSuperTip(this.dtFechaHasta, null);
			this.dtFechaHasta.TabIndex = 87;
			this.dtFechaHasta.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// dtFechaDesde
			// 
			this.dtFechaDesde.AllowMonthSelection = true;
			this.dtFechaDesde.AllowWeekSelection = true;
			this.dtFechaDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaDesde.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsComision, "Com_CtaComision.FecDesde"));
			dateButton2.Caption = "Today";
			this.dtFechaDesde.DateButtons.Add(dateButton2);
			this.dtFechaDesde.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.dtFechaDesde.Format = "dd/MMM/yyyy";
			this.dtFechaDesde.Location = new System.Drawing.Point(624, 8);
			this.dtFechaDesde.Name = "dtFechaDesde";
			this.dtFechaDesde.NonAutoSizeHeight = 23;
			this.dtFechaDesde.Size = new System.Drawing.Size(112, 21);
			this.dtFechaDesde.SpinButtonsVisible = true;
			this.toolTipController1.SetSuperTip(this.dtFechaDesde, null);
			this.dtFechaDesde.TabIndex = 86;
			this.dtFechaDesde.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(568, 32);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.toolTipController1.SetSuperTip(this.label2, null);
			this.label2.TabIndex = 85;
			this.label2.Text = "Hasta:";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(568, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.toolTipController1.SetSuperTip(this.label1, null);
			this.label1.TabIndex = 84;
			this.label1.Text = "Desde:";
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 10);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(43, 16);
			this.toolTipController1.SetSuperTip(this.lblCliente, null);
			this.lblCliente.TabIndex = 89;
			this.lblCliente.Text = "Código:";
			// 
			// txtCodigo
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance50;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsComision, "Com_CtaComision.Codigo"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(96, 8);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(112, 21);
			this.toolTipController1.SetSuperTip(this.txtCodigo, null);
			this.txtCodigo.TabIndex = 88;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 34);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(67, 16);
			this.toolTipController1.SetSuperTip(this.label3, null);
			this.label3.TabIndex = 91;
			this.label3.Text = "Descripción:";
			// 
			// txtDescripcion
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance51;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsComision, "Com_CtaComision.Descripcion"));
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(96, 32);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(352, 21);
			this.toolTipController1.SetSuperTip(this.txtDescripcion, null);
			this.txtDescripcion.TabIndex = 90;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(776, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.toolTipController1.SetSuperTip(this.label4, null);
			this.label4.TabIndex = 128;
			this.label4.Text = "Estado:";
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1214, 396);
			this.toolTipController1.SetSuperTip(this.ultraTabSharedControlsPage1, null);
			// 
			// utcTransaccion
			// 
			this.utcTransaccion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcTransaccion.Appearance = appearance52;
			this.utcTransaccion.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl1);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl2);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl3);
			this.utcTransaccion.Controls.Add(this.pgEntrega);
			this.utcTransaccion.Location = new System.Drawing.Point(8, 64);
			this.utcTransaccion.Name = "utcTransaccion";
			this.utcTransaccion.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcTransaccion.Size = new System.Drawing.Size(1216, 417);
			this.utcTransaccion.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.toolTipController1.SetSuperTip(this.utcTransaccion, null);
			this.utcTransaccion.TabIndex = 129;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Metas";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Rol x Ejecutivo";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Comisiones";
			ultraTab4.TabPage = this.pgEntrega;
			ultraTab4.Text = "Evaluadores y Roles";
			this.utcTransaccion.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																														 ultraTab1,
																																														 ultraTab2,
																																														 ultraTab3,
																																														 ultraTab4});
			this.utcTransaccion.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// cmbEstado
			// 
			this.cmbEstado.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsComision, "Com_CtaComision.Status"));
			this.cmbEstado.Items.AddRange(new object[] {
																									 "ABIERTO",
																									 "CERRADO"});
			this.cmbEstado.Location = new System.Drawing.Point(840, 8);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(121, 21);
			this.toolTipController1.SetSuperTip(this.cmbEstado, null);
			this.cmbEstado.TabIndex = 132;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Com_CtaComision";
			this.barraDato1.DataNombreId = "idComision";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsComision;
			this.barraDato1.DataTabla = "Com_CtaComision";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(24, 487);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.toolTipController1.SetSuperTip(this.barraDato1, null);
			this.barraDato1.TabIndex = 148;
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
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
			// txtComision
			// 
			this.txtComision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComision.Appearance = appearance53;
			this.txtComision.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsComision, "Com_CtaComision.idComision"));
			this.txtComision.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComision.FormatString = "";
			this.txtComision.Location = new System.Drawing.Point(1160, 488);
			this.txtComision.Name = "txtComision";
			this.txtComision.PromptChar = ' ';
			this.txtComision.ReadOnly = true;
			this.txtComision.Size = new System.Drawing.Size(64, 21);
			this.toolTipController1.SetSuperTip(this.txtComision, null);
			this.txtComision.TabIndex = 150;
			// 
			// btnTransforma
			// 
			this.btnTransforma.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance54.Image = ((object)(resources.GetObject("appearance54.Image")));
			this.btnTransforma.Appearance = appearance54;
			this.btnTransforma.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnTransforma.Location = new System.Drawing.Point(280, 487);
			this.btnTransforma.Name = "btnTransforma";
			this.btnTransforma.Size = new System.Drawing.Size(24, 25);
			this.toolTipController1.SetSuperTip(this.btnTransforma, null);
			this.btnTransforma.TabIndex = 151;
			this.btnTransforma.Tag = "";
			this.toolTip.SetToolTip(this.btnTransforma, "Duplicar COMISIONES");
			this.btnTransforma.Click += new System.EventHandler(this.btnTransforma_Click);
			// 
			// toolTipController1
			// 
			this.toolTipController1.ToolTipType = DevExpress.Utils.ToolTipType.SuperTip;
			// 
			// Point_ConfComision
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1232, 517);
			this.Controls.Add(this.btnTransforma);
			this.Controls.Add(this.txtComision);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.utcTransaccion);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.dtFechaHasta);
			this.Controls.Add(this.dtFechaDesde);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Name = "Point_ConfComision";
			this.toolTipController1.SetSuperTip(this, null);
			this.Text = "Configuración de Comisiones";
			this.Load += new System.EventHandler(this.Point_ConfComision_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdMetas)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbRolVen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEjecutivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRolEjecutivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTPCalificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdComisiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.pgEntrega.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbEvalua5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEvalua4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEvalua3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEvalua2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdSueldos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEvalua1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcTransaccion)).EndInit();
			this.utcTransaccion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComision)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		C1.Data.C1DataRow drComisiones;
		C1.Data.C1DataRow drSeteoF;
    CultureInfo us = new CultureInfo("en-US");
		bool bEliminaAlValidar = false;

		private void cdsComision_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsComision.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void Point_ConfComision_Load(object sender, System.EventArgs e)
		{
			barraDato1.HabilitaControles(false);
			barraDato1.PosUltima();
			CargarDatos();
			HabilitaTab();
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
		}

		private void CargarDatos()
		{
			
			string idCodComisn = (string)this.txtCodigo.Value;
			if(idCodComisn!=null)
			{
				string sQry = string.Format("Exec Point_Comisiones '{0}'", idCodComisn);
				SqlDataReader dReader = Funcion.rEscalarSQL(cdsComision,sQry,true);
				dReader.Read();
				
				int idComsion = (int)dReader.GetValue(0);
				dReader.Close();

				# region Datos Comisiones
				string stExec = string.Format("Exec POINT_Porcentaje_Comision {0}",idComsion);
				grdComisiones.DataSource = Funcion.dvProcedimiento(cdsComision, stExec);
				# endregion Datos Comisiones

				# region Datos Metas
				string stQExec = string.Format("Exec Point_Metas_Comision {0}",idComsion);
				grdMetas.DataSource = Funcion.dvProcedimiento(cdsComision, stQExec);
				# endregion Datos Metas
				
				# region Datos RolesxEjectvo Grid
				string stRExec = string.Format("Exec Point_Roles_Comision {0}",idComsion);
				grdRolEjecutivo.DataSource = Funcion.dvProcedimiento(cdsComision, stRExec);
				# endregion Datos RolesxEjectvo
				
				/*# region Datos Evaluadores
				string stEExec = string.Format("Exec Point_EvaluaRoles_Comision {0}",idComsion);
				grdSueldos.DataSource = Funcion.dvProcedimiento(cdsComision, stEExec);
				# endregion Datos Evaluadores*/

				# region Datos ComboEvaluador
				string stEv1 = "Exec Point_PersonalVentas_Comision 1";
				cmbEvalua1.DataSource = Funcion.dvProcedimiento(cdsComision, stEv1);
				string stEv2 = "Exec Point_PersonalVentas_Comision 1";
				cmbEvalua2.DataSource = Funcion.dvProcedimiento(cdsComision, stEv2);
				string stEv3 = "Exec Point_PersonalVentas_Comision 1";
				cmbEvalua3.DataSource = Funcion.dvProcedimiento(cdsComision, stEv3);
				string stEv4 = "Exec Point_PersonalVentas_Comision 1";
				cmbEvalua4.DataSource = Funcion.dvProcedimiento(cdsComision, stEv4);
				string stEv5 = "Exec Point_PersonalVentas_Comision 1";
				cmbEvalua5.DataSource = Funcion.dvProcedimiento(cdsComision, stEv5);
				# endregion Datos ComboEvaluador

				# region Datos ComboBodegas
				string stCmBodg = string.Format("Exec Point_Bodega_Comision {0}",1);
				cmbLocal.DataSource = Funcion.dvProcedimiento(cdsComision, stCmBodg);
				# endregion Datos ComboBodegas

				# region Datos ComboRolEjectivo
				string stCmRole = string.Format("Exec Point_PersonalVentas_Comision {0}",1);
				cmbEjecutivo.DataSource = Funcion.dvProcedimiento(cdsComision, stCmRole);
				# endregion Datos ComboRolEjectivo

				# region Datos Combo Cargos Rol
				string stCmbCargo = string.Format("Exec Point_PersonalVentas_Comision {0}",2);
				cmbCargo.DataSource = Funcion.dvProcedimiento(cdsComision, stCmbCargo);
				# endregion Datos Combo Cargos Rol


			}
		
		}

		private void HabilitaTab()
		{
			foreach(Infragistics.Win.UltraWinTabControl.UltraTab ctr in utcTransaccion.Tabs)
				foreach(Control cn in ctr.TabPage.Controls)
					barraDato1.HabilitaControl(cn, barraDato1.bEditar);
			//utcTransaccion.Enabled = barraDato1.bEditar;
		}

		private void grdComisiones_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			/* Cambiar el color de una línea*/
			if ((int)e.Row.Cells["idTipoVenta"].Value == 1)
			{
				e.Row.Appearance.BackColor = Color.BurlyWood;
				e.Row.Appearance.BackColor2 = Color.White;
				e.Row.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			}
			if ((int)e.Row.Cells["idTipoVenta"].Value == 2)
			{
				e.Row.Appearance.BackColor = Color.Aqua;
				e.Row.Appearance.BackColor2 = Color.White;
				e.Row.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			}
			if ((int)e.Row.Cells["idTipoVenta"].Value == 3)
			{
				e.Row.Appearance.BackColor = Color.Beige;
				e.Row.Appearance.BackColor2 = Color.White;
				e.Row.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			}
			if ((int)e.Row.Cells["idTipoVenta"].Value == 4)
			{
				e.Row.Appearance.BackColor = Color.AliceBlue;
				e.Row.Appearance.BackColor2 = Color.White;
				e.Row.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			}
			if ((int)e.Row.Cells["idTipoVenta"].Value == 5)
			{
				e.Row.Appearance.BackColor = Color.Aquamarine;
				e.Row.Appearance.BackColor2 = Color.White;
				e.Row.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			}
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drComisiones = cdsComision.TableViews["Com_CtaComision"].Rows[0];
				CargarDatos();
			}
			catch{};

		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			this.barraDato1.BorraRegistro();

		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
			HabilitaTab();
			/* Limpiar GRIDS de las pantallas */
			# region Datos Metas
			string stQExec = string.Format("Exec Point_Metas_Comision {0}",0);
			grdMetas.DataSource = Funcion.dvProcedimiento(cdsComision, stQExec);
			# endregion Datos Metas
			# region Datos RolesxEjectvo Grid
			string stRExec = string.Format("Exec Point_Roles_Comision {0}",0);
			grdRolEjecutivo.DataSource = Funcion.dvProcedimiento(cdsComision, stRExec);
			# endregion Datos RolesxEjectvo
			# region Datos Comisiones
			string stExec = string.Format("Exec POINT_Porcentaje_comision {0}",0);
			grdComisiones.DataSource = Funcion.dvProcedimiento(cdsComision, stExec);
			# endregion Datos Comisiones

		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
			HabilitaTab();

		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
			HabilitaTab();

		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			this.barraDato1.GrabaRegistro();
			HabilitaTab();
			//Cursor = Cursors.WaitCursor;
			#region Graba Metas
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdMetas.Rows.All)
			{
				#region Borra Filas que no tienen articulos ingresados
				if (dr.Cells["Meta"].Value == System.DBNull.Value || dr.Cells["Tipo"].Value == System.DBNull.Value || dr.Cells["Cant_Vend"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}
				#endregion Borra Filas que no tienen articulos ingresados
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drCuotas in grdMetas.Rows.All)
			{
				#region Datos Cuotas
				int idComis = (int)this.txtComision.Value;
        int idReg = 0;
				if (drCuotas.Cells["idRegistro"].Value!=System.DBNull.Value) idReg = (int)drCuotas.Cells["idRegistro"].Value;
        string idTp = "";
				if (drCuotas.Cells["Tipo"].Value!=System.DBNull.Value) idTp = drCuotas.Cells["Tipo"].Value.ToString();
				string CodUL = "";
				if (drCuotas.Cells["CodigoLU"].Value!=System.DBNull.Value) CodUL = drCuotas.Cells["CodigoLU"].Value.ToString();
				double nVendrs = 0;
				if (drCuotas.Cells["Cant_Vend"].Value!=System.DBNull.Value) nVendrs = (double)drCuotas.Cells["Cant_Vend"].Value;
				double Mta = 0;
				if (drCuotas.Cells["Meta"].Value!=System.DBNull.Value) Mta = (double)drCuotas.Cells["Meta"].Value;
				double pContd = 0;
				if (drCuotas.Cells["PctContado"].Value!=System.DBNull.Value) pContd = (double)drCuotas.Cells["PctContado"].Value;
				double cContd = 0;
				if (drCuotas.Cells["CuotaContado"].Value!=System.DBNull.Value) cContd = (double)drCuotas.Cells["CuotaContado"].Value;
				double pCred = 0;
				if (drCuotas.Cells["PctCredito"].Value!=System.DBNull.Value) pCred = (double)drCuotas.Cells["PctCredito"].Value;
				double cCred = 0;
				if (drCuotas.Cells["CuotaCredito"].Value!=System.DBNull.Value) cCred = (double)drCuotas.Cells["CuotaCredito"].Value;
				double cZens = 0;
				if (drCuotas.Cells["CuotaZense"].Value!=System.DBNull.Value) cZens = (double)drCuotas.Cells["CuotaZense"].Value;
				double ndias = 0;
				if (drCuotas.Cells["Dias"].Value!=System.DBNull.Value) ndias = (double)drCuotas.Cells["Dias"].Value;
				string Trimestr = this.txtCodigo.Value.ToString();
				#endregion Datos Cuotas

				string sqlTran = string.Format("Exec Point_Comisiones_IA {0},{1},'{2}','{3}','{4}',{5},{6},{7},{8},{9},{10},{11},{12}",
					idComis,idReg,Trimestr,CodUL,idTp,Mta,pContd,pCred,cContd,cCred,nVendrs,ndias,cZens);
				Funcion.EjecutaSQL(cdsSeteoF, sqlTran, true);
			}
			#endregion Graba Metas

			#region Graba Roles
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drGR in grdRolEjecutivo.Rows.All)
			{
				#region Borra Filas que no tienen articulos ingresados
				if (drGR.Cells["idPersonal"].Value == System.DBNull.Value || drGR.Cells["idCargo"].Value == System.DBNull.Value || drGR.Cells["idRolVentas"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					drGR.Delete();
				}
				#endregion Borra Filas que no tienen articulos ingresados
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drRoles in grdRolEjecutivo.Rows.All)
			{
				#region Datos Cuotas
				int idComis = (int)this.txtComision.Value;
				int idRol = 0;
				if (drRoles.Cells["idRol"].Value!=System.DBNull.Value) idRol = (int)drRoles.Cells["idRol"].Value;
				int idPers = 0;
				if (drRoles.Cells["idPersonal"].Value!=System.DBNull.Value) idPers = (int)drRoles.Cells["idPersonal"].Value;
				string idCarg = "";
				if (drRoles.Cells["idCargo"].Value!=System.DBNull.Value) idCarg = drRoles.Cells["idCargo"].Value.ToString();
				DateTime FIni = DateTime.Today;
				if (drRoles.Cells["FechaInicia"].Value!=System.DBNull.Value) FIni = (DateTime)drRoles.Cells["FechaInicia"].Value;
				DateTime FFin = DateTime.Today;
				if (drRoles.Cells["FechaFin"].Value!=System.DBNull.Value) FFin = (DateTime)drRoles.Cells["FechaFin"].Value;
				string Tiends = "";
				if (drRoles.Cells["Tiendas"].Value!=System.DBNull.Value) Tiends = drRoles.Cells["Tiendas"].Value.ToString();
				int TpRolVen = 0;
				if (drRoles.Cells["idRolVentas"].Value!=System.DBNull.Value) TpRolVen = (int)drRoles.Cells["idRolVentas"].Value;
				int nDias = 0;
				if (drRoles.Cells["Cant_Dias"].Value!=System.DBNull.Value) nDias = (int)drRoles.Cells["Cant_Dias"].Value;
				int EdoEjecutivo = 0;
				if (drRoles.Cells["Estado"].Value!=System.DBNull.Value) EdoEjecutivo = (int)drRoles.Cells["Estado"].Value;
				DateTime FecSalida = DateTime.Today;
				if (drRoles.Cells["FechaSalida"].Value!=System.DBNull.Value) 
				  FecSalida = (DateTime)drRoles.Cells["FechaSalida"].Value;
				else
					FecSalida = Convert.ToDateTime("01/01/1900");
				string Trimestr = this.txtCodigo.Value.ToString();
				#endregion Datos Cuotas

				string sqlTran = string.Format("Exec Point_Comisiones_IARoles {0},{1},'{2}',{3},'{4}',{5},{6},'{7}','{8}',{9},{10},'{11}'",
					idComis,idRol,Trimestr,idPers,Tiends,idCarg,nDias,FIni.ToString("yyyyMMdd"),FFin.ToString("yyyyMMdd"),TpRolVen,EdoEjecutivo,FecSalida.ToString("yyyyMMdd"));
				Funcion.EjecutaSQL(cdsSeteoF, sqlTran, true);
			}
			#endregion Graba Roles

			#region Graba Porcentajes
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drGP in grdComisiones.Rows.All)
			{
				#region Borra Filas que no tienen articulos ingresados
				if (drGP.Cells["idTipoVenta"].Value == System.DBNull.Value || drGP.Cells["idTipoCalificacion"].Value == System.DBNull.Value || drGP.Cells["VMinimo"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					drGP.Delete();
				}
				#endregion Borra Filas que no tienen articulos ingresados
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drPorcntjs in grdComisiones.Rows.All)
			{
				#region Datos Porcentajes
				int idComis = (int)this.txtComision.Value;
				int idPorc = 0;
				if (drPorcntjs.Cells["idPorcentaje"].Value!=System.DBNull.Value)
					idPorc = (int)drPorcntjs.Cells["idPorcentaje"].Value;
				int idTPV = 0;
				if (drPorcntjs.Cells["idTipoVenta"].Value!=System.DBNull.Value)
					idTPV = (int)drPorcntjs.Cells["idTipoVenta"].Value;
				int idTPC = 0;
				if (drPorcntjs.Cells["idTipoCalificacion"].Value!=System.DBNull.Value)
					idTPC = (int)drPorcntjs.Cells["idTipoCalificacion"].Value;
				double ValMi = 0;
				if (drPorcntjs.Cells["VMinimo"].Value!=System.DBNull.Value)
					ValMi = (double)drPorcntjs.Cells["VMinimo"].Value;
				double ValMx = 0;
				if (drPorcntjs.Cells["VMaximo"].Value!=System.DBNull.Value)
					ValMx = (double)drPorcntjs.Cells["VMaximo"].Value;
				double ValGV = 0;
				if (drPorcntjs.Cells["VGrteVentas"].Value!=System.DBNull.Value)
					ValGV = (double)drPorcntjs.Cells["VGrteVentas"].Value;
				double ValJfV = 0;
				if (drPorcntjs.Cells["VJefeVentas"].Value!=System.DBNull.Value)
					ValJfV = (double)drPorcntjs.Cells["VJefeVentas"].Value;
				double ValSup = 0;
				if (drPorcntjs.Cells["VSupervisor"].Value!=System.DBNull.Value)
					ValSup = (double)drPorcntjs.Cells["VSupervisor"].Value;
				double ValJfTA = 0;
				if (drPorcntjs.Cells["VJefeTiendaA"].Value!=System.DBNull.Value)
					ValJfTA = (double)drPorcntjs.Cells["VJefeTiendaA"].Value;
				double ValJfTB = 0;
				if (drPorcntjs.Cells["VJefeTiendaB"].Value!=System.DBNull.Value)
					ValJfTB = (double)drPorcntjs.Cells["VJefeTiendaB"].Value;
				double ValJfTC = 0;
				if (drPorcntjs.Cells["VJefeTiendaC"].Value!=System.DBNull.Value)
					ValJfTC = (double)drPorcntjs.Cells["VJefeTiendaC"].Value;
				double ValVn = 0;
				if (drPorcntjs.Cells["VVendedor"].Value!=System.DBNull.Value)
					ValVn = (double)drPorcntjs.Cells["VVendedor"].Value;
				string Trimestr = this.txtCodigo.Value.ToString();
				#endregion Datos Porcentajes

				string sqlTran = string.Format("Exec Point_Comisiones_IAPorcntj {0},{1},'{2}',{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}",
					idComis,idPorc,Trimestr,idTPV,idTPC,ValMi,ValMx,ValGV,ValJfV,ValSup,ValJfTA,ValJfTB,ValJfTC,ValVn);
				Funcion.EjecutaSQL(cdsSeteoF, sqlTran, true);
			}
			#endregion Graba Roles
		}

		private void grdMetas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			

		}

		private void grdMetas_AfterCellUpdate_1(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Creacion de Variables
			double dValMeta = 0;
			double dValorContado = 0;
			double dValorCredito = 0;
			double dPCredito = 0;
			double dPContado = 0;
			#endregion Creacion de Variables

			#region Tomar Valores Grilla
			if (e.Cell.Row.Cells["Meta"].Value!=System.DBNull.Value) dValMeta = (double) e.Cell.Row.Cells["Meta"].Value;
			if (e.Cell.Row.Cells["CuotaContado"].Value!=System.DBNull.Value) dValorContado = (double)e.Cell.Row.Cells["CuotaContado"].Value;
			if (e.Cell.Row.Cells["CuotaCredito"].Value!=System.DBNull.Value) dValorCredito = (double)e.Cell.Row.Cells["CuotaCredito"].Value;
			#endregion Tomar Valores Grilla

			if (e.Cell.Column.ToString() == "TIPO")
			{
				if ((int)e.Cell.Row.Cells["Tipo"].Value == 1) e.Cell.Row.Cells["Cant_Vend"].Value = 1;
			}

			if(dValorContado>dValMeta)
			{
				MessageBox.Show("El Valor de CONTADO no puede ser mayor al valor de la meta..?", "Aviso..!!");
				e.Cell.Row.Cells["CuotaContado"].Value = 0;
				e.Cell.Row.Cells["CuotaCredito"].Value = 0;
				return;
			}
			else
			{
				if(dValMeta+dValorContado>0)
					e.Cell.Row.Cells["CuotaCredito"].Value = dValMeta-dValorContado;

				if(dValMeta+dValorContado+dValorCredito>0)
				{
					//dPContado = Math.Round(dValorContado/dValMeta*100,6);
					//dPCredito = Math.Round(dValorCredito/dValMeta*100,6);
					dPContado = dValorContado/dValMeta*100;
					dPCredito = dValorCredito/dValMeta*100;

					e.Cell.Row.Cells["PctContado"].Value = dPContado;
					e.Cell.Row.Cells["PctCredito"].Value = dPCredito;
				}
			}
		}

		private void grdRolEjecutivo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			/* Llamar Funcion F3 */
			int idDLPer = (int)grdRolEjecutivo.ActiveRow.Cells["idCodven"].Value;
			if (e.KeyValue == (int)Keys.F3)
			{
				POINT_Selecciona_Bodegas miConfComi = new POINT_Selecciona_Bodegas(idDLPer);
				miConfComi.ShowDialog();
			}
		}

		private void grdComisiones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMonddec = "00.000000";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stCrs = "000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			
			#region Decimales para la cantidad, peso y Pedido
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["VMinimo"].Format = stFormato;
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["VMaximo"].Format = stFormato;
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["VGrteVentas"].Format = stFormato;
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["VJefeVentas"].Format = stFormato;
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["VSupervisor"].Format = stFormato;
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["VJefeTiendaA"].Format = stFormato;
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["VJefeTiendaB"].Format = stFormato;
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["VJefeTiendaC"].Format = stFormato;
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["VVendedor"].Format = stFormato;
			#endregion

		}

		private void grdMetas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			//int x = (int)drSeteoF["decCantidad"];
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMonddec = "00.000000";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stCrs = "000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			
			#region Decimales para la cantidad, peso y Pedido
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decCantidad"]);
			e.Layout.Bands[0].Columns["Dias"].Format = stFormato;
			
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["Meta"].Format = stFormato;
			
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["CuotaContado"].Format = stFormato;
			
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["CuotaCredito"].Format = stFormato;
			
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["CuotaZense"].Format = stFormato;
			
			stFormato = stMonddec + stCrs.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["PctContado"].Format = stFormato;
			
			stFormato = stMonddec + stCrs.Substring(0, (int) drSeteoF["decPrecio"]);
			e.Layout.Bands[0].Columns["PctCredito"].Format = stFormato;
			#endregion

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void grdMetas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (grdMetas.ActiveCell == null) return;
			try
			{
				//if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					#region Page Down Abre lista desplegable
					case (int) Keys.PageDown:
						if(grdMetas.ActiveCell.DroppedDown == false)
							grdMetas.ActiveCell.DroppedDown = true;
						break;
					#endregion Page Down Abre lista desplegable
					#region Flechas
					case (int) Keys.Up:
						if (grdMetas.ActiveCell.DroppedDown) return;

						grdMetas.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdMetas.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						grdMetas.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						grdMetas.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdMetas.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						grdMetas.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (grdMetas.ActiveCell.DroppedDown) return;
			
						grdMetas.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdMetas.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						grdMetas.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					#endregion Flechas
					#region Enter
					case (int) Keys.Enter:
						grdMetas.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdMetas.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						grdMetas.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(grdMetas.ActiveCell.DroppedDown == false)
							grdMetas.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
					#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void grdRolEjecutivo_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (grdRolEjecutivo.ActiveCell == null) return;
			try
			{
				//if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
						#region Page Down Abre lista desplegable
					case (int) Keys.PageDown:
						if(grdRolEjecutivo.ActiveCell.DroppedDown == false)
							grdRolEjecutivo.ActiveCell.DroppedDown = true;
						break;
						#endregion Page Down Abre lista desplegable
						#region Flechas
					case (int) Keys.Up:
						if (grdRolEjecutivo.ActiveCell.DroppedDown) return;

						grdRolEjecutivo.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdRolEjecutivo.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						grdRolEjecutivo.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						grdRolEjecutivo.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdRolEjecutivo.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						grdRolEjecutivo.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (grdRolEjecutivo.ActiveCell.DroppedDown) return;
			
						grdRolEjecutivo.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdRolEjecutivo.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						grdRolEjecutivo.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
						#endregion Flechas
						#region Enter
					case (int) Keys.Enter:
						grdRolEjecutivo.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdRolEjecutivo.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						grdRolEjecutivo.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(grdRolEjecutivo.ActiveCell.DroppedDown == false)
							grdRolEjecutivo.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void grdRolEjecutivo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMonddec = "00.000000";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stCrs = "000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			
			#region Decimales para la cantidad, peso y Pedido
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decCantidad"]);
			e.Layout.Bands[0].Columns["Cant_Dias"].Format = stFormato;
			#endregion

		}

		private void grdComisiones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (grdComisiones.ActiveCell == null) return;
			try
			{
				//if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
						#region Page Down Abre lista desplegable
					case (int) Keys.PageDown:
						if(grdComisiones.ActiveCell.DroppedDown == false)
							grdComisiones.ActiveCell.DroppedDown = true;
						break;
						#endregion Page Down Abre lista desplegable
						#region Flechas
					case (int) Keys.Up:
						if (grdComisiones.ActiveCell.DroppedDown) return;

						grdComisiones.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdComisiones.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						grdComisiones.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						grdComisiones.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdComisiones.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						grdComisiones.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (grdComisiones.ActiveCell.DroppedDown) return;
			
						grdComisiones.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdComisiones.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						grdComisiones.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
						#endregion Flechas
						#region Enter
					case (int) Keys.Enter:
						grdComisiones.PerformAction(UltraGridAction.ExitEditMode, false, false);
						grdComisiones.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						grdComisiones.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(grdComisiones.ActiveCell.DroppedDown == false)
							grdComisiones.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void btnTransforma_Click(object sender, System.EventArgs e)
		{
			// Duplicar comisiones
			#region Datos idComision
			int idComis = 0;
			if (this.txtComision.Value!=System.DBNull.Value) idComis = (int)this.txtComision.Value;
			#endregion Datos idComision

			if (DialogResult.No != MessageBox.Show("Está Seguro de Duplicar COMISIONES.?","Confirmación",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2))
			{
				string sqlDupli = string.Format("Exec Point_Duplica_Comisiones {0}",idComis);
				Funcion.EjecutaSQL(cdsSeteoF, sqlDupli, true);
				MessageBox.Show("Dato Duplicado, proceda a modificar la información duplicada..!!");
			}

		}

		private void grdMetas_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			/* Borrar registro*/
			//int idCuota = (int)e.Rows[1].Cells["idRegistro"].Value;
			int idCuota = 0;
			
			if (DialogResult.No != MessageBox.Show("Está Seguro de Eliminar Registro.?","Confirmación",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2))
			{
				for (int i=0; i<e.Rows.Length; i++)
				{
					idCuota = (int)e.Rows[i].Cells["idRegistro"].Value;
					string sqlDupli = string.Format("Exec Elimina_Registro_Cuota {0},{1}",idCuota,1);
					Funcion.EjecutaSQL(cdsSeteoF, sqlDupli, true);
					MessageBox.Show("Registro eliminado..!!");
				}
        e.DisplayPromptMsg = false;
			}
			else 
				e.Cancel = true;
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void grdMetas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.grdMetas.ActiveCell.Column.Key.ToString().ToUpper() == "CANT_VEND") Funcion.CamposNumericos(sender, e);
				 //if ((int)this.grdMetas.ActiveRow.Cells["Tipo"].Value == 1 && this.grdMetas.ActiveCell.Column.Key.ToString().ToUpper() == "CANT_VEND") e.Handled = true;
		}

		private void grdRolEjecutivo_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			/* Borrar registro*/
			//int idCuota = (int)e.Rows[1].Cells["idRegistro"].Value;
			int idRlEje = 0;
			
			if (DialogResult.No != MessageBox.Show("Está Seguro de Eliminar Registro.?","Confirmación",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2))
			{
				for (int i=0; i<e.Rows.Length; i++)
				{
					idRlEje = (int)e.Rows[i].Cells["idRol"].Value;
					string sqlEliRE = string.Format("Exec Elimina_Registro_Cuota {0},{1}",idRlEje,2);
					Funcion.EjecutaSQL(cdsSeteoF, sqlEliRE, true);
					MessageBox.Show("Registro eliminado..!!");
				}
				e.DisplayPromptMsg = false;
			}
			else 
				e.Cancel = true;
		}

		private void grdComisiones_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			/* Borrar registro*/
			//int idCuota = (int)e.Rows[1].Cells["idRegistro"].Value;
			int idPrcn = 0;
			
			if (DialogResult.No != MessageBox.Show("Está Seguro de Eliminar Registro.?","Confirmación",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2))
			{
				for (int i=0; i<e.Rows.Length; i++)
				{
					idPrcn = (int)e.Rows[i].Cells["idPorcentaje"].Value;
					string sqlEliRE = string.Format("Exec Elimina_Registro_Cuota {0},{1}",idPrcn,3);
					Funcion.EjecutaSQL(cdsSeteoF, sqlEliRE, true);
					MessageBox.Show("Registro eliminado..!!");
				}
				e.DisplayPromptMsg = false;
			}
			else 
				e.Cancel = true;

		}

	}
}
