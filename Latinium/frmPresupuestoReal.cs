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
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPresupuestoReal.
	/// </summary>
	public class frmPresupuestoReal : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnProyeccion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosTotal;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadAgosto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label16;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadAbril;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosSeptiembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIngresosEnero;
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAño;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVentas;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVentasPrecios;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVentasMargen;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridOtrosIngresos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPEnero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPAgosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPJulio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPDiciembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPNoviembre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPOctubre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGPSeptiembre;
		private C1.Data.C1DataSet c1DataSet1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor6;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP22;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP23;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP21;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP19;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP23T;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP22T;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP21T;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP20T;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP19T;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGP18T;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadTotalP;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidadTotalUN;
		private System.Windows.Forms.Label lblprueba;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinGrid.UltraGrid UGridUb;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtdETJunio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtdETMayo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtdETAbril;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtdETMarzo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtdETFebrero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtdETEnero;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPresupuestoReal()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPresupuestoReal));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2018");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2019");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2020");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2021");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2022");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2023");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2018");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2019");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2020");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2021");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2022");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2023");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Margen");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2018");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2019");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2020");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2021");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2022");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2023");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo_1");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo_2");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo_3");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo_4");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo_5");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo_6");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo_2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo_3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo_4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo_5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo_6");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2018");
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2019");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2020");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2021");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2022");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2023");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2019", 0);
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2020", 1);
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2021", 2);
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2022", 3);
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2023", 4);
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 10, true);
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2018", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2018", 4, true);
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2019", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2019", 5, true);
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2020", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2020", 6, true);
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2021", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2021", 7, true);
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2022", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2022", 8, true);
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2023", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2023", 9, true);
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2019", 0, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2019", 0, false);
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2020", 1, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2020", 1, false);
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2021", 2, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2021", 2, false);
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2022", 3, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2022", 3, false);
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2023", 4, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2023", 4, false);
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Porcentaje", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Porcentaje", 3, true);
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2018");
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2019");
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2020");
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2021");
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2022");
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2023");
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Margen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2019", 0);
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2020", 1);
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2021", 2);
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2022", 3);
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2023", 4);
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 8, true);
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2018", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2018", 2, true);
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2019", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2019", 3, true);
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2020", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2020", 4, true);
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2021", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2021", 5, true);
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2022", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2022", 6, true);
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2023", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2023", 7, true);
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2019", 0, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2019", 0, false);
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2020", 1, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2020", 1, false);
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2021", 2, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2021", 2, false);
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2022", 3, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2022", 3, false);
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2023", 4, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2023", 4, false);
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Porcentaje", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Porcentaje", 1, true);
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2018");
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2019");
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2020");
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2021");
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2022");
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2023");
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2019", 0);
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2020", 1);
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2021", 2);
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2022", 3);
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2023", 4);
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings27 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 11, true);
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings28 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2018", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2018", 5, true);
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings29 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2019", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2019", 6, true);
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings30 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2020", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2020", 7, true);
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings31 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2021", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2021", 8, true);
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings32 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2022", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2022", 9, true);
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings33 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2023", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2023", 10, true);
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings34 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Porcentaje", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Porcentaje", 3, true);
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings35 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2019", 0, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2019", 0, false);
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings36 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2020", 1, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2020", 1, false);
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings37 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2021", 2, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2021", 2, false);
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings38 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2022", 3, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2022", 3, false);
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings39 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2023", 4, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2023", 4, false);
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance155 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance156 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance157 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.Appearance appearance158 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAgrupa");
			Infragistics.Win.Appearance appearance159 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance160 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2018");
			Infragistics.Win.Appearance appearance161 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2019");
			Infragistics.Win.Appearance appearance162 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2020");
			Infragistics.Win.Appearance appearance163 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2021");
			Infragistics.Win.Appearance appearance164 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2022");
			Infragistics.Win.Appearance appearance165 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2023");
			Infragistics.Win.Appearance appearance166 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance167 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance168 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2019", 0);
			Infragistics.Win.Appearance appearance169 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2020", 1);
			Infragistics.Win.Appearance appearance170 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2021", 2);
			Infragistics.Win.Appearance appearance171 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2022", 3);
			Infragistics.Win.Appearance appearance172 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2023", 4);
			Infragistics.Win.Appearance appearance173 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings40 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2018", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2018", 4, true);
			Infragistics.Win.Appearance appearance174 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings41 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2019", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2019", 5, true);
			Infragistics.Win.Appearance appearance175 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings42 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2020", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2020", 6, true);
			Infragistics.Win.Appearance appearance176 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings43 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2021", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2021", 7, true);
			Infragistics.Win.Appearance appearance177 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings44 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2022", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2022", 8, true);
			Infragistics.Win.Appearance appearance178 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings45 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2023", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2023", 9, true);
			Infragistics.Win.Appearance appearance179 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings46 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 10, true);
			Infragistics.Win.Appearance appearance180 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings47 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Porcentaje", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Porcentaje", 11, true);
			Infragistics.Win.Appearance appearance181 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings48 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2019", 0, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2019", 0, false);
			Infragistics.Win.Appearance appearance182 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings49 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2020", 1, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2020", 1, false);
			Infragistics.Win.Appearance appearance183 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings50 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2021", 2, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2021", 2, false);
			Infragistics.Win.Appearance appearance184 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings51 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2022", 3, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2022", 3, false);
			Infragistics.Win.Appearance appearance185 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings52 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2023", 4, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2023", 4, false);
			Infragistics.Win.Appearance appearance186 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance187 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance188 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance189 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance190 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance191 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance192 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance193 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAgrupa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2018");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2019");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2020");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2021");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2022");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2023");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Utilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2018");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2019");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2020");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2021");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2022");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("P_2023");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("por_2019");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("por_2020");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("por_2021");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("por_2022");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("por_2023");
			Infragistics.Win.Appearance appearance226 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Utilidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance227 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2018");
			Infragistics.Win.Appearance appearance228 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2019");
			Infragistics.Win.Appearance appearance229 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2020");
			Infragistics.Win.Appearance appearance230 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2021");
			Infragistics.Win.Appearance appearance231 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2022");
			Infragistics.Win.Appearance appearance232 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2023");
			Infragistics.Win.Appearance appearance233 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance234 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2019");
			Infragistics.Win.Appearance appearance235 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2020");
			Infragistics.Win.Appearance appearance236 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2021");
			Infragistics.Win.Appearance appearance237 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2022");
			Infragistics.Win.Appearance appearance238 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2023");
			Infragistics.Win.Appearance appearance239 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings53 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2018", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2018", 2, true);
			Infragistics.Win.Appearance appearance240 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings54 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2019", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2019", 3, true);
			Infragistics.Win.Appearance appearance241 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings55 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2021", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2021", 5, true);
			Infragistics.Win.Appearance appearance242 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings56 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2022", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2022", 6, true);
			Infragistics.Win.Appearance appearance243 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings57 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2023", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2023", 7, true);
			Infragistics.Win.Appearance appearance244 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings58 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 8, true);
			Infragistics.Win.Appearance appearance245 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings59 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Porcentaje", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Porcentaje", 1, true);
			Infragistics.Win.Appearance appearance246 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings60 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2019", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2019", 9, true);
			Infragistics.Win.Appearance appearance247 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings61 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2020", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2020", 10, true);
			Infragistics.Win.Appearance appearance248 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings62 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2021", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2021", 11, true);
			Infragistics.Win.Appearance appearance249 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings63 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2022", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2022", 12, true);
			Infragistics.Win.Appearance appearance250 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings64 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "por_2023", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "por_2023", 13, true);
			Infragistics.Win.Appearance appearance251 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings65 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "P_2020", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "P_2020", 4, true);
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Utilidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2018");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2019");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2020");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2021");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2022");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("P_2023");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2019");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2020");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2021");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2022");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("por_2023");
			Infragistics.Win.Appearance appearance267 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance268 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance269 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance270 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance271 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance272 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance273 = new Infragistics.Win.Appearance();
			this.btnProyeccion = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.txtUtilidadTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtUtilidadNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label17 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label12 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtUtilidadMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtUtilidadEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIngresosEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.txtAño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.uGridVentas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridVentasPrecios = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridVentasMargen = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridOtrosIngresos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label18 = new System.Windows.Forms.Label();
			this.txtGPEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGPFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGPAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGPMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGPAgosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGPJulio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGPJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGPMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGPDiciembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGPNoviembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGPOctubre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGPSeptiembre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.c1DataSet1 = new C1.Data.C1DataSet();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor6 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGP22 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGP23 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGP20 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGP21 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGP18 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGP19 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUtilidadTotalP = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.txtGP23T = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGP22T = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGP21T = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGP20T = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGP19T = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGP18T = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.txtUtilidadTotalUN = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblprueba = new System.Windows.Forms.Label();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.UGridUb = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label22 = new System.Windows.Forms.Label();
			this.txtdETJunio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtdETMayo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtdETAbril = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtdETMarzo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtdETFebrero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtdETEnero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasPrecios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasMargen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOtrosIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPAgosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPJulio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPDiciembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPNoviembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPOctubre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPSeptiembre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataSet1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP22)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP23)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP21)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP18)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP19)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadTotalP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP23T)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP22T)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP21T)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP20T)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP19T)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP18T)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadTotalUN)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.UGridUb)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETJunio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETMayo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETAbril)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETMarzo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETFebrero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETEnero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// btnProyeccion
			// 
			this.btnProyeccion.CausesValidation = false;
			this.btnProyeccion.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnProyeccion.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnProyeccion.Image = ((System.Drawing.Image)(resources.GetObject("btnProyeccion.Image")));
			this.btnProyeccion.Location = new System.Drawing.Point(24, 504);
			this.btnProyeccion.Name = "btnProyeccion";
			this.btnProyeccion.Size = new System.Drawing.Size(30, 23);
			this.btnProyeccion.TabIndex = 868;
			this.btnProyeccion.Visible = false;
			this.btnProyeccion.Click += new System.EventHandler(this.btnProyeccion_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.Decimal);
			ultraDataColumn2.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn3.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
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
																																 ultraDataColumn14});
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Firebrick;
			this.label3.Location = new System.Drawing.Point(96, 560);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(80, 16);
			this.label3.TabIndex = 855;
			this.label3.Text = "2018";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtUtilidadTotal
			// 
			this.txtUtilidadTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.FontData.BoldAsString = "True";
			appearance1.ForeColor = System.Drawing.Color.Firebrick;
			appearance1.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadTotal.Appearance = appearance1;
			this.txtUtilidadTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadTotal.Enabled = false;
			this.txtUtilidadTotal.FormatString = "#,##0.00";
			this.txtUtilidadTotal.Location = new System.Drawing.Point(624, 640);
			this.txtUtilidadTotal.Name = "txtUtilidadTotal";
			this.txtUtilidadTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadTotal.PromptChar = ' ';
			this.txtUtilidadTotal.Size = new System.Drawing.Size(100, 21);
			this.txtUtilidadTotal.TabIndex = 854;
			// 
			// txtGastosTotal
			// 
			this.txtGastosTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance2.FontData.BoldAsString = "True";
			appearance2.ForeColor = System.Drawing.Color.Firebrick;
			appearance2.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtGastosTotal.Appearance = appearance2;
			this.txtGastosTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosTotal.Enabled = false;
			this.txtGastosTotal.FormatString = "#,##0.00";
			this.txtGastosTotal.Location = new System.Drawing.Point(624, 608);
			this.txtGastosTotal.Name = "txtGastosTotal";
			this.txtGastosTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosTotal.PromptChar = ' ';
			this.txtGastosTotal.Size = new System.Drawing.Size(100, 21);
			this.txtGastosTotal.TabIndex = 853;
			// 
			// txtIngresosTotal
			// 
			this.txtIngresosTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance3.FontData.BoldAsString = "True";
			appearance3.ForeColor = System.Drawing.Color.Firebrick;
			appearance3.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtIngresosTotal.Appearance = appearance3;
			this.txtIngresosTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosTotal.Enabled = false;
			this.txtIngresosTotal.FormatString = "#,##0.00";
			this.txtIngresosTotal.Location = new System.Drawing.Point(624, 584);
			this.txtIngresosTotal.Name = "txtIngresosTotal";
			this.txtIngresosTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosTotal.PromptChar = ' ';
			this.txtIngresosTotal.Size = new System.Drawing.Size(100, 21);
			this.txtIngresosTotal.TabIndex = 852;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Firebrick;
			this.label2.Location = new System.Drawing.Point(56, 640);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(31, 16);
			this.label2.TabIndex = 851;
			this.label2.Text = "(=) U";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(56, 608);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(28, 16);
			this.label1.TabIndex = 850;
			this.label1.Text = "(-) G";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(56, 584);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(26, 16);
			this.label5.TabIndex = 849;
			this.label5.Text = "(+) I";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtUtilidadNoviembre
			// 
			this.txtUtilidadNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance4.ForeColor = System.Drawing.Color.Firebrick;
			appearance4.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadNoviembre.Appearance = appearance4;
			this.txtUtilidadNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadNoviembre.Enabled = false;
			this.txtUtilidadNoviembre.FormatString = "#,##0.00";
			this.txtUtilidadNoviembre.Location = new System.Drawing.Point(976, 640);
			this.txtUtilidadNoviembre.Name = "txtUtilidadNoviembre";
			this.txtUtilidadNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadNoviembre.PromptChar = ' ';
			this.txtUtilidadNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadNoviembre.TabIndex = 848;
			this.txtUtilidadNoviembre.Visible = false;
			// 
			// txtUtilidadDiciembre
			// 
			this.txtUtilidadDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance5.ForeColor = System.Drawing.Color.Firebrick;
			appearance5.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadDiciembre.Appearance = appearance5;
			this.txtUtilidadDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadDiciembre.Enabled = false;
			this.txtUtilidadDiciembre.FormatString = "#,##0.00";
			this.txtUtilidadDiciembre.Location = new System.Drawing.Point(1064, 640);
			this.txtUtilidadDiciembre.Name = "txtUtilidadDiciembre";
			this.txtUtilidadDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadDiciembre.PromptChar = ' ';
			this.txtUtilidadDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadDiciembre.TabIndex = 847;
			this.txtUtilidadDiciembre.Visible = false;
			// 
			// txtUtilidadSeptiembre
			// 
			this.txtUtilidadSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance6.ForeColor = System.Drawing.Color.Firebrick;
			appearance6.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadSeptiembre.Appearance = appearance6;
			this.txtUtilidadSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadSeptiembre.Enabled = false;
			this.txtUtilidadSeptiembre.FormatString = "#,##0.00";
			this.txtUtilidadSeptiembre.Location = new System.Drawing.Point(800, 640);
			this.txtUtilidadSeptiembre.Name = "txtUtilidadSeptiembre";
			this.txtUtilidadSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadSeptiembre.PromptChar = ' ';
			this.txtUtilidadSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadSeptiembre.TabIndex = 846;
			this.txtUtilidadSeptiembre.Visible = false;
			// 
			// txtUtilidadOctubre
			// 
			this.txtUtilidadOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance7.ForeColor = System.Drawing.Color.Firebrick;
			appearance7.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadOctubre.Appearance = appearance7;
			this.txtUtilidadOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadOctubre.Enabled = false;
			this.txtUtilidadOctubre.FormatString = "#,##0.00";
			this.txtUtilidadOctubre.Location = new System.Drawing.Point(888, 640);
			this.txtUtilidadOctubre.Name = "txtUtilidadOctubre";
			this.txtUtilidadOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadOctubre.PromptChar = ' ';
			this.txtUtilidadOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadOctubre.TabIndex = 845;
			this.txtUtilidadOctubre.Visible = false;
			// 
			// txtUtilidadJulio
			// 
			this.txtUtilidadJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance8.ForeColor = System.Drawing.Color.Firebrick;
			appearance8.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadJulio.Appearance = appearance8;
			this.txtUtilidadJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadJulio.Enabled = false;
			this.txtUtilidadJulio.FormatString = "#,##0.00";
			this.txtUtilidadJulio.Location = new System.Drawing.Point(1152, 640);
			this.txtUtilidadJulio.Name = "txtUtilidadJulio";
			this.txtUtilidadJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadJulio.PromptChar = ' ';
			this.txtUtilidadJulio.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadJulio.TabIndex = 844;
			this.txtUtilidadJulio.Visible = false;
			// 
			// txtUtilidadAgosto
			// 
			this.txtUtilidadAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance9.ForeColor = System.Drawing.Color.Firebrick;
			appearance9.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadAgosto.Appearance = appearance9;
			this.txtUtilidadAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadAgosto.Enabled = false;
			this.txtUtilidadAgosto.FormatString = "#,##0.00";
			this.txtUtilidadAgosto.Location = new System.Drawing.Point(712, 640);
			this.txtUtilidadAgosto.Name = "txtUtilidadAgosto";
			this.txtUtilidadAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadAgosto.PromptChar = ' ';
			this.txtUtilidadAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadAgosto.TabIndex = 843;
			this.txtUtilidadAgosto.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn25});
			// 
			// label17
			// 
			this.label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label17.Location = new System.Drawing.Point(8, 688);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(1240, 1);
			this.label17.TabIndex = 869;
			this.label17.Click += new System.EventHandler(this.label17_Click);
			// 
			// label16
			// 
			this.label16.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.ForeColor = System.Drawing.Color.Firebrick;
			this.label16.Location = new System.Drawing.Point(616, 560);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(100, 16);
			this.label16.TabIndex = 867;
			this.label16.Text = "TOTAL";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.ForeColor = System.Drawing.Color.Firebrick;
			this.label12.Location = new System.Drawing.Point(1064, 560);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(80, 16);
			this.label12.TabIndex = 866;
			this.label12.Text = "DICIEMBRE";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label12.Visible = false;
			// 
			// label13
			// 
			this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.ForeColor = System.Drawing.Color.Firebrick;
			this.label13.Location = new System.Drawing.Point(976, 560);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(80, 16);
			this.label13.TabIndex = 865;
			this.label13.Text = "NOVIEMBRE";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label13.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35});
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.ForeColor = System.Drawing.Color.Firebrick;
			this.label14.Location = new System.Drawing.Point(888, 560);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(80, 16);
			this.label14.TabIndex = 864;
			this.label14.Text = "OCTUBRE";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label14.Visible = false;
			// 
			// label15
			// 
			this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.ForeColor = System.Drawing.Color.Firebrick;
			this.label15.Location = new System.Drawing.Point(800, 560);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(80, 16);
			this.label15.TabIndex = 863;
			this.label15.Text = "SEPTIEMBRE";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label15.Visible = false;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.Firebrick;
			this.label8.Location = new System.Drawing.Point(712, 560);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 862;
			this.label8.Text = "AGOSTO";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label8.Visible = false;
			// 
			// label9
			// 
			this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.Firebrick;
			this.label9.Location = new System.Drawing.Point(1160, 560);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 16);
			this.label9.TabIndex = 861;
			this.label9.Text = "JULIO";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label9.Visible = false;
			// 
			// label10
			// 
			this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.ForeColor = System.Drawing.Color.Firebrick;
			this.label10.Location = new System.Drawing.Point(536, 560);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(80, 16);
			this.label10.TabIndex = 860;
			this.label10.Text = "2023";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label10.Click += new System.EventHandler(this.label10_Click);
			// 
			// label11
			// 
			this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.ForeColor = System.Drawing.Color.Firebrick;
			this.label11.Location = new System.Drawing.Point(448, 560);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(80, 16);
			this.label11.TabIndex = 859;
			this.label11.Text = "2022";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.ForeColor = System.Drawing.Color.Firebrick;
			this.label7.Location = new System.Drawing.Point(360, 560);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 16);
			this.label7.TabIndex = 858;
			this.label7.Text = "2021";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.ForeColor = System.Drawing.Color.Firebrick;
			this.label6.Location = new System.Drawing.Point(272, 560);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 16);
			this.label6.TabIndex = 857;
			this.label6.Text = "2020";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.ForeColor = System.Drawing.Color.Firebrick;
			this.label4.Location = new System.Drawing.Point(184, 560);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 16);
			this.label4.TabIndex = 856;
			this.label4.Text = "2019";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtUtilidadMayo
			// 
			this.txtUtilidadMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance10.ForeColor = System.Drawing.Color.Firebrick;
			appearance10.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadMayo.Appearance = appearance10;
			this.txtUtilidadMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadMayo.Enabled = false;
			this.txtUtilidadMayo.FormatString = "#,##0.00";
			this.txtUtilidadMayo.Location = new System.Drawing.Point(448, 640);
			this.txtUtilidadMayo.Name = "txtUtilidadMayo";
			this.txtUtilidadMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadMayo.PromptChar = ' ';
			this.txtUtilidadMayo.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadMayo.TabIndex = 842;
			// 
			// txtUtilidadJunio
			// 
			this.txtUtilidadJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance11.ForeColor = System.Drawing.Color.Firebrick;
			appearance11.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadJunio.Appearance = appearance11;
			this.txtUtilidadJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadJunio.Enabled = false;
			this.txtUtilidadJunio.FormatString = "#,##0.00";
			this.txtUtilidadJunio.Location = new System.Drawing.Point(536, 640);
			this.txtUtilidadJunio.Name = "txtUtilidadJunio";
			this.txtUtilidadJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadJunio.PromptChar = ' ';
			this.txtUtilidadJunio.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadJunio.TabIndex = 841;
			// 
			// txtUtilidadMarzo
			// 
			this.txtUtilidadMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance12.ForeColor = System.Drawing.Color.Firebrick;
			appearance12.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadMarzo.Appearance = appearance12;
			this.txtUtilidadMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadMarzo.Enabled = false;
			this.txtUtilidadMarzo.FormatString = "#,##0.00";
			this.txtUtilidadMarzo.Location = new System.Drawing.Point(272, 640);
			this.txtUtilidadMarzo.Name = "txtUtilidadMarzo";
			this.txtUtilidadMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadMarzo.PromptChar = ' ';
			this.txtUtilidadMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadMarzo.TabIndex = 840;
			// 
			// txtUtilidadAbril
			// 
			this.txtUtilidadAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance13.ForeColor = System.Drawing.Color.Firebrick;
			appearance13.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadAbril.Appearance = appearance13;
			this.txtUtilidadAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadAbril.Enabled = false;
			this.txtUtilidadAbril.FormatString = "#,##0.00";
			this.txtUtilidadAbril.Location = new System.Drawing.Point(360, 640);
			this.txtUtilidadAbril.Name = "txtUtilidadAbril";
			this.txtUtilidadAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadAbril.PromptChar = ' ';
			this.txtUtilidadAbril.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadAbril.TabIndex = 839;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn47.DataType = typeof(System.Decimal);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47});
			// 
			// ultraDataSource4
			// 
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(int);
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn52.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn53.DataType = typeof(System.Decimal);
			ultraDataColumn53.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn54.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn55.DataType = typeof(System.Decimal);
			ultraDataColumn55.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn56.DataType = typeof(System.Decimal);
			ultraDataColumn56.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn57.DataType = typeof(System.Decimal);
			ultraDataColumn57.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn58.DataType = typeof(System.Decimal);
			ultraDataColumn58.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn59.DataType = typeof(System.Decimal);
			ultraDataColumn59.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn60.DataType = typeof(System.Decimal);
			ultraDataColumn60.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn61.DataType = typeof(System.Decimal);
			ultraDataColumn61.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn62.DataType = typeof(System.Decimal);
			ultraDataColumn62.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn63.DataType = typeof(System.Decimal);
			ultraDataColumn63.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn64.DataType = typeof(System.Decimal);
			ultraDataColumn64.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn65.DataType = typeof(System.Decimal);
			ultraDataColumn65.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54,
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58,
																																 ultraDataColumn59,
																																 ultraDataColumn60,
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63,
																																 ultraDataColumn64,
																																 ultraDataColumn65});
			// 
			// txtUtilidadEnero
			// 
			this.txtUtilidadEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance14.ForeColor = System.Drawing.Color.Firebrick;
			appearance14.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadEnero.Appearance = appearance14;
			this.txtUtilidadEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadEnero.Enabled = false;
			this.txtUtilidadEnero.FormatString = "#,##0.00";
			this.txtUtilidadEnero.Location = new System.Drawing.Point(96, 640);
			this.txtUtilidadEnero.Name = "txtUtilidadEnero";
			this.txtUtilidadEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadEnero.PromptChar = ' ';
			this.txtUtilidadEnero.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadEnero.TabIndex = 838;
			// 
			// txtUtilidadFebrero
			// 
			this.txtUtilidadFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance15.ForeColor = System.Drawing.Color.Firebrick;
			appearance15.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadFebrero.Appearance = appearance15;
			this.txtUtilidadFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadFebrero.Enabled = false;
			this.txtUtilidadFebrero.FormatString = "#,##0.00";
			this.txtUtilidadFebrero.Location = new System.Drawing.Point(184, 640);
			this.txtUtilidadFebrero.Name = "txtUtilidadFebrero";
			this.txtUtilidadFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadFebrero.PromptChar = ' ';
			this.txtUtilidadFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtUtilidadFebrero.TabIndex = 837;
			// 
			// txtGastosNoviembre
			// 
			this.txtGastosNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosNoviembre.Appearance = appearance16;
			this.txtGastosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosNoviembre.Enabled = false;
			this.txtGastosNoviembre.FormatString = "#,##0.00";
			this.txtGastosNoviembre.Location = new System.Drawing.Point(976, 608);
			this.txtGastosNoviembre.Name = "txtGastosNoviembre";
			this.txtGastosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosNoviembre.PromptChar = ' ';
			this.txtGastosNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosNoviembre.TabIndex = 836;
			this.txtGastosNoviembre.Visible = false;
			// 
			// txtGastosDiciembre
			// 
			this.txtGastosDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosDiciembre.Appearance = appearance17;
			this.txtGastosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosDiciembre.Enabled = false;
			this.txtGastosDiciembre.FormatString = "#,##0.00";
			this.txtGastosDiciembre.Location = new System.Drawing.Point(1064, 608);
			this.txtGastosDiciembre.Name = "txtGastosDiciembre";
			this.txtGastosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosDiciembre.PromptChar = ' ';
			this.txtGastosDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosDiciembre.TabIndex = 835;
			this.txtGastosDiciembre.Visible = false;
			// 
			// txtIngresosDiciembre
			// 
			this.txtIngresosDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosDiciembre.Appearance = appearance18;
			this.txtIngresosDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosDiciembre.Enabled = false;
			this.txtIngresosDiciembre.FormatString = "#,##0.00";
			this.txtIngresosDiciembre.Location = new System.Drawing.Point(1064, 584);
			this.txtIngresosDiciembre.Name = "txtIngresosDiciembre";
			this.txtIngresosDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosDiciembre.PromptChar = ' ';
			this.txtIngresosDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosDiciembre.TabIndex = 834;
			this.txtIngresosDiciembre.Visible = false;
			// 
			// txtIngresosNoviembre
			// 
			this.txtIngresosNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosNoviembre.Appearance = appearance19;
			this.txtIngresosNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosNoviembre.Enabled = false;
			this.txtIngresosNoviembre.FormatString = "#,##0.00";
			this.txtIngresosNoviembre.Location = new System.Drawing.Point(976, 584);
			this.txtIngresosNoviembre.Name = "txtIngresosNoviembre";
			this.txtIngresosNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosNoviembre.PromptChar = ' ';
			this.txtIngresosNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosNoviembre.TabIndex = 833;
			this.txtIngresosNoviembre.Visible = false;
			// 
			// txtGastosSeptiembre
			// 
			this.txtGastosSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosSeptiembre.Appearance = appearance20;
			this.txtGastosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosSeptiembre.Enabled = false;
			this.txtGastosSeptiembre.FormatString = "#,##0.00";
			this.txtGastosSeptiembre.Location = new System.Drawing.Point(800, 608);
			this.txtGastosSeptiembre.Name = "txtGastosSeptiembre";
			this.txtGastosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosSeptiembre.PromptChar = ' ';
			this.txtGastosSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosSeptiembre.TabIndex = 832;
			this.txtGastosSeptiembre.Visible = false;
			// 
			// txtGastosOctubre
			// 
			this.txtGastosOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosOctubre.Appearance = appearance21;
			this.txtGastosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosOctubre.Enabled = false;
			this.txtGastosOctubre.FormatString = "#,##0.00";
			this.txtGastosOctubre.Location = new System.Drawing.Point(888, 608);
			this.txtGastosOctubre.Name = "txtGastosOctubre";
			this.txtGastosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosOctubre.PromptChar = ' ';
			this.txtGastosOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtGastosOctubre.TabIndex = 831;
			this.txtGastosOctubre.Visible = false;
			// 
			// txtIngresosOctubre
			// 
			this.txtIngresosOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosOctubre.Appearance = appearance22;
			this.txtIngresosOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosOctubre.Enabled = false;
			this.txtIngresosOctubre.FormatString = "#,##0.00";
			this.txtIngresosOctubre.Location = new System.Drawing.Point(888, 584);
			this.txtIngresosOctubre.Name = "txtIngresosOctubre";
			this.txtIngresosOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosOctubre.PromptChar = ' ';
			this.txtIngresosOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosOctubre.TabIndex = 830;
			this.txtIngresosOctubre.Visible = false;
			// 
			// txtIngresosSeptiembre
			// 
			this.txtIngresosSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosSeptiembre.Appearance = appearance23;
			this.txtIngresosSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosSeptiembre.Enabled = false;
			this.txtIngresosSeptiembre.FormatString = "#,##0.00";
			this.txtIngresosSeptiembre.Location = new System.Drawing.Point(800, 584);
			this.txtIngresosSeptiembre.Name = "txtIngresosSeptiembre";
			this.txtIngresosSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosSeptiembre.PromptChar = ' ';
			this.txtIngresosSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosSeptiembre.TabIndex = 829;
			this.txtIngresosSeptiembre.Visible = false;
			// 
			// txtGastosJulio
			// 
			this.txtGastosJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosJulio.Appearance = appearance24;
			this.txtGastosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosJulio.Enabled = false;
			this.txtGastosJulio.FormatString = "#,##0.00";
			this.txtGastosJulio.Location = new System.Drawing.Point(1152, 608);
			this.txtGastosJulio.Name = "txtGastosJulio";
			this.txtGastosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosJulio.PromptChar = ' ';
			this.txtGastosJulio.Size = new System.Drawing.Size(80, 21);
			this.txtGastosJulio.TabIndex = 828;
			this.txtGastosJulio.Visible = false;
			// 
			// txtGastosAgosto
			// 
			this.txtGastosAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosAgosto.Appearance = appearance25;
			this.txtGastosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosAgosto.Enabled = false;
			this.txtGastosAgosto.FormatString = "#,##0.00";
			this.txtGastosAgosto.Location = new System.Drawing.Point(712, 608);
			this.txtGastosAgosto.Name = "txtGastosAgosto";
			this.txtGastosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosAgosto.PromptChar = ' ';
			this.txtGastosAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtGastosAgosto.TabIndex = 827;
			this.txtGastosAgosto.Visible = false;
			// 
			// txtIngresosAgosto
			// 
			this.txtIngresosAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosAgosto.Appearance = appearance26;
			this.txtIngresosAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosAgosto.Enabled = false;
			this.txtIngresosAgosto.FormatString = "#,##0.00";
			this.txtIngresosAgosto.Location = new System.Drawing.Point(712, 584);
			this.txtIngresosAgosto.Name = "txtIngresosAgosto";
			this.txtIngresosAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosAgosto.PromptChar = ' ';
			this.txtIngresosAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosAgosto.TabIndex = 826;
			this.txtIngresosAgosto.Visible = false;
			// 
			// txtIngresosJulio
			// 
			this.txtIngresosJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosJulio.Appearance = appearance27;
			this.txtIngresosJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosJulio.Enabled = false;
			this.txtIngresosJulio.FormatString = "#,##0.00";
			this.txtIngresosJulio.Location = new System.Drawing.Point(1152, 584);
			this.txtIngresosJulio.Name = "txtIngresosJulio";
			this.txtIngresosJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosJulio.PromptChar = ' ';
			this.txtIngresosJulio.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosJulio.TabIndex = 825;
			this.txtIngresosJulio.Visible = false;
			// 
			// txtGastosMayo
			// 
			this.txtGastosMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosMayo.Appearance = appearance28;
			this.txtGastosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosMayo.Enabled = false;
			this.txtGastosMayo.FormatString = "#,##0.00";
			this.txtGastosMayo.Location = new System.Drawing.Point(448, 608);
			this.txtGastosMayo.Name = "txtGastosMayo";
			this.txtGastosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosMayo.PromptChar = ' ';
			this.txtGastosMayo.Size = new System.Drawing.Size(80, 21);
			this.txtGastosMayo.TabIndex = 824;
			// 
			// txtGastosJunio
			// 
			this.txtGastosJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosJunio.Appearance = appearance29;
			this.txtGastosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosJunio.Enabled = false;
			this.txtGastosJunio.FormatString = "#,##0.00";
			this.txtGastosJunio.Location = new System.Drawing.Point(536, 608);
			this.txtGastosJunio.Name = "txtGastosJunio";
			this.txtGastosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosJunio.PromptChar = ' ';
			this.txtGastosJunio.Size = new System.Drawing.Size(80, 21);
			this.txtGastosJunio.TabIndex = 823;
			// 
			// txtIngresosJunio
			// 
			this.txtIngresosJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosJunio.Appearance = appearance30;
			this.txtIngresosJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosJunio.Enabled = false;
			this.txtIngresosJunio.FormatString = "#,##0.00";
			this.txtIngresosJunio.Location = new System.Drawing.Point(536, 584);
			this.txtIngresosJunio.Name = "txtIngresosJunio";
			this.txtIngresosJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosJunio.PromptChar = ' ';
			this.txtIngresosJunio.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosJunio.TabIndex = 822;
			// 
			// txtIngresosMayo
			// 
			this.txtIngresosMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosMayo.Appearance = appearance31;
			this.txtIngresosMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosMayo.Enabled = false;
			this.txtIngresosMayo.FormatString = "#,##0.00";
			this.txtIngresosMayo.Location = new System.Drawing.Point(448, 584);
			this.txtIngresosMayo.Name = "txtIngresosMayo";
			this.txtIngresosMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosMayo.PromptChar = ' ';
			this.txtIngresosMayo.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosMayo.TabIndex = 821;
			// 
			// txtGastosMarzo
			// 
			this.txtGastosMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosMarzo.Appearance = appearance32;
			this.txtGastosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosMarzo.Enabled = false;
			this.txtGastosMarzo.FormatString = "#,##0.00";
			this.txtGastosMarzo.Location = new System.Drawing.Point(272, 608);
			this.txtGastosMarzo.Name = "txtGastosMarzo";
			this.txtGastosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosMarzo.PromptChar = ' ';
			this.txtGastosMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtGastosMarzo.TabIndex = 820;
			// 
			// txtGastosAbril
			// 
			this.txtGastosAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosAbril.Appearance = appearance33;
			this.txtGastosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosAbril.Enabled = false;
			this.txtGastosAbril.FormatString = "#,##0.00";
			this.txtGastosAbril.Location = new System.Drawing.Point(360, 608);
			this.txtGastosAbril.Name = "txtGastosAbril";
			this.txtGastosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosAbril.PromptChar = ' ';
			this.txtGastosAbril.Size = new System.Drawing.Size(80, 21);
			this.txtGastosAbril.TabIndex = 819;
			// 
			// txtIngresosAbril
			// 
			this.txtIngresosAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosAbril.Appearance = appearance34;
			this.txtIngresosAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosAbril.Enabled = false;
			this.txtIngresosAbril.FormatString = "#,##0.00";
			this.txtIngresosAbril.Location = new System.Drawing.Point(360, 584);
			this.txtIngresosAbril.Name = "txtIngresosAbril";
			this.txtIngresosAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosAbril.PromptChar = ' ';
			this.txtIngresosAbril.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosAbril.TabIndex = 818;
			// 
			// txtIngresosMarzo
			// 
			this.txtIngresosMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosMarzo.Appearance = appearance35;
			this.txtIngresosMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosMarzo.Enabled = false;
			this.txtIngresosMarzo.FormatString = "#,##0.00";
			this.txtIngresosMarzo.Location = new System.Drawing.Point(272, 584);
			this.txtIngresosMarzo.Name = "txtIngresosMarzo";
			this.txtIngresosMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosMarzo.PromptChar = ' ';
			this.txtIngresosMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosMarzo.TabIndex = 817;
			// 
			// txtGastosEnero
			// 
			this.txtGastosEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosEnero.Appearance = appearance36;
			this.txtGastosEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosEnero.Enabled = false;
			this.txtGastosEnero.FormatString = "#,##0.00";
			this.txtGastosEnero.Location = new System.Drawing.Point(96, 608);
			this.txtGastosEnero.Name = "txtGastosEnero";
			this.txtGastosEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosEnero.PromptChar = ' ';
			this.txtGastosEnero.Size = new System.Drawing.Size(80, 21);
			this.txtGastosEnero.TabIndex = 816;
			// 
			// txtGastosFebrero
			// 
			this.txtGastosFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastosFebrero.Appearance = appearance37;
			this.txtGastosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastosFebrero.Enabled = false;
			this.txtGastosFebrero.FormatString = "#,##0.00";
			this.txtGastosFebrero.Location = new System.Drawing.Point(184, 608);
			this.txtGastosFebrero.Name = "txtGastosFebrero";
			this.txtGastosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosFebrero.PromptChar = ' ';
			this.txtGastosFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtGastosFebrero.TabIndex = 815;
			// 
			// txtIngresosFebrero
			// 
			this.txtIngresosFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosFebrero.Appearance = appearance38;
			this.txtIngresosFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosFebrero.Enabled = false;
			this.txtIngresosFebrero.FormatString = "#,##0.00";
			this.txtIngresosFebrero.Location = new System.Drawing.Point(184, 584);
			this.txtIngresosFebrero.Name = "txtIngresosFebrero";
			this.txtIngresosFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosFebrero.PromptChar = ' ';
			this.txtIngresosFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosFebrero.TabIndex = 814;
			// 
			// txtIngresosEnero
			// 
			this.txtIngresosEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresosEnero.Appearance = appearance39;
			this.txtIngresosEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresosEnero.Enabled = false;
			this.txtIngresosEnero.FormatString = "#,##0.00";
			this.txtIngresosEnero.Location = new System.Drawing.Point(96, 584);
			this.txtIngresosEnero.Name = "txtIngresosEnero";
			this.txtIngresosEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIngresosEnero.PromptChar = ' ';
			this.txtIngresosEnero.Size = new System.Drawing.Size(80, 21);
			this.txtIngresosEnero.TabIndex = 813;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.Location = new System.Drawing.Point(24, 568);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(30, 23);
			this.btnNuevo.TabIndex = 810;
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Visible = false;
			// 
			// txtAño
			// 
			this.txtAño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAño.Location = new System.Drawing.Point(48, 16);
			this.txtAño.MinValue = 2017;
			this.txtAño.Name = "txtAño";
			this.txtAño.PromptChar = ' ';
			this.txtAño.Size = new System.Drawing.Size(88, 21);
			this.txtAño.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAño.TabIndex = 806;
			this.txtAño.Visible = false;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.Location = new System.Drawing.Point(8, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(30, 23);
			this.btnVer.TabIndex = 807;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// uGridVentas
			// 
			this.uGridVentas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVentas.DataSource = this.ultraDataSource7;
			appearance40.BackColor = System.Drawing.Color.White;
			this.uGridVentas.DisplayLayout.Appearance = appearance40;
			this.uGridVentas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn1.CellAppearance = appearance41;
			ultraGridColumn1.Format = "#,##0.00";
			ultraGridColumn1.Header.Caption = "2018";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 178;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.CellAppearance = appearance42;
			ultraGridColumn2.Format = "#,##0.00";
			ultraGridColumn2.Header.Caption = "2019";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 173;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance43;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.Caption = "2020";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 173;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance44;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "2021";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 173;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance45;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "2022";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 173;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance46;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "2023";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 173;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand1.SummaryFooterCaption = "TOTALES";
			this.uGridVentas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance47.ForeColor = System.Drawing.Color.Black;
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentas.DisplayLayout.Override.ActiveRowAppearance = appearance47;
			this.uGridVentas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVentas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance48.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentas.DisplayLayout.Override.CardAreaAppearance = appearance48;
			appearance49.ForeColor = System.Drawing.Color.Black;
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentas.DisplayLayout.Override.CellAppearance = appearance49;
			this.uGridVentas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance50.FontData.BoldAsString = "True";
			appearance50.FontData.Name = "Arial";
			appearance50.FontData.SizeInPoints = 8F;
			appearance50.ForeColor = System.Drawing.Color.White;
			appearance50.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentas.DisplayLayout.Override.HeaderAppearance = appearance50;
			appearance51.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance51.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.RowAlternateAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.RowSelectorAppearance = appearance52;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentas.DisplayLayout.Override.SelectedRowAppearance = appearance53;
			this.uGridVentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentas.Location = new System.Drawing.Point(152, 5);
			this.uGridVentas.Name = "uGridVentas";
			this.uGridVentas.Size = new System.Drawing.Size(1081, 40);
			this.uGridVentas.TabIndex = 805;
			this.uGridVentas.Visible = false;
			this.uGridVentas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVentas_KeyDown);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn66.DataType = typeof(System.Decimal);
			ultraDataColumn67.DataType = typeof(System.Decimal);
			ultraDataColumn68.DataType = typeof(System.Decimal);
			ultraDataColumn69.DataType = typeof(System.Decimal);
			ultraDataColumn70.DataType = typeof(System.Decimal);
			ultraDataColumn71.DataType = typeof(System.Decimal);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn66,
																																 ultraDataColumn67,
																																 ultraDataColumn68,
																																 ultraDataColumn69,
																																 ultraDataColumn70,
																																 ultraDataColumn71});
			// 
			// uGridVentasPrecios
			// 
			this.uGridVentasPrecios.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVentasPrecios.DataSource = this.ultraDataSource2;
			appearance54.BackColor = System.Drawing.Color.White;
			this.uGridVentasPrecios.DisplayLayout.Appearance = appearance54;
			this.uGridVentasPrecios.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Año";
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 137;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Venta";
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 229;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance55;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "%";
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 91;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance56;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Width = 85;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance57;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Width = 79;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance58;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Width = 67;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance59;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Width = 67;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance60;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Width = 54;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance61;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 13;
			ultraGridColumn16.Width = 55;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance62;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 15;
			ultraGridColumn17.Width = 147;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance63;
			ultraGridColumn18.DataType = typeof(System.Decimal);
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "%";
			ultraGridColumn18.Header.VisiblePosition = 6;
			ultraGridColumn18.Width = 52;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance64;
			ultraGridColumn19.DataType = typeof(System.Decimal);
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "%";
			ultraGridColumn19.Header.VisiblePosition = 8;
			ultraGridColumn19.Width = 55;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance65;
			ultraGridColumn20.DataType = typeof(System.Decimal);
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "%";
			ultraGridColumn20.Header.VisiblePosition = 10;
			ultraGridColumn20.Width = 62;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance66;
			ultraGridColumn21.DataType = typeof(System.Decimal);
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "%";
			ultraGridColumn21.Header.VisiblePosition = 12;
			ultraGridColumn21.Width = 69;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance67;
			ultraGridColumn22.DataType = typeof(System.Decimal);
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "%";
			ultraGridColumn22.Header.VisiblePosition = 14;
			ultraGridColumn22.Width = 76;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn22});
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance68;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance69.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance69;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance70;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance71;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance72;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance73.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance73;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance74;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance75;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance76;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance77;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance78;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance79;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance80;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
			ultraGridBand2.SummaryFooterCaption = "TOTALES";
			this.uGridVentasPrecios.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance81.ForeColor = System.Drawing.Color.Black;
			appearance81.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasPrecios.DisplayLayout.Override.ActiveRowAppearance = appearance81;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentasPrecios.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance82.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentasPrecios.DisplayLayout.Override.CardAreaAppearance = appearance82;
			appearance83.ForeColor = System.Drawing.Color.Black;
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasPrecios.DisplayLayout.Override.CellAppearance = appearance83;
			this.uGridVentasPrecios.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance84.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance84.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance84.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance84.FontData.BoldAsString = "True";
			appearance84.FontData.Name = "Arial";
			appearance84.FontData.SizeInPoints = 8F;
			appearance84.ForeColor = System.Drawing.Color.White;
			appearance84.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentasPrecios.DisplayLayout.Override.HeaderAppearance = appearance84;
			appearance85.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance85.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance85.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.RowAlternateAppearance = appearance85;
			appearance86.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance86.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance86.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.RowSelectorAppearance = appearance86;
			appearance87.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance87.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance87.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasPrecios.DisplayLayout.Override.SelectedRowAppearance = appearance87;
			this.uGridVentasPrecios.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentasPrecios.Location = new System.Drawing.Point(8, 45);
			this.uGridVentasPrecios.Name = "uGridVentasPrecios";
			this.uGridVentasPrecios.Size = new System.Drawing.Size(1226, 83);
			this.uGridVentasPrecios.TabIndex = 808;
			this.uGridVentasPrecios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVentasPrecios_KeyDown);
			this.uGridVentasPrecios.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVentasPrecios_InitializeLayout);
			// 
			// uGridVentasMargen
			// 
			this.uGridVentasMargen.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVentasMargen.DataSource = this.ultraDataSource3;
			appearance88.BackColor = System.Drawing.Color.White;
			this.uGridVentasMargen.DisplayLayout.Appearance = appearance88;
			this.uGridVentasMargen.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.Caption = "Año";
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 137;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance89.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance89;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "%";
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn24.Width = 148;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance90.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance90;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 2;
			ultraGridColumn25.Width = 58;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance91.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance91;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 4;
			ultraGridColumn26.Width = 58;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance92.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance92;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 6;
			ultraGridColumn27.Width = 58;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance93.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance93;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 8;
			ultraGridColumn28.Width = 58;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance94;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 10;
			ultraGridColumn29.Width = 58;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance95;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 12;
			ultraGridColumn30.Width = 60;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance96.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance96;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 14;
			ultraGridColumn31.Width = 163;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.Caption = "Costo de Venta";
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Width = 172;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance97.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance97;
			ultraGridColumn33.DataType = typeof(System.Decimal);
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.Caption = "%";
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn33.Width = 63;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance98.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance98;
			ultraGridColumn34.DataType = typeof(System.Decimal);
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.Caption = "%";
			ultraGridColumn34.Header.VisiblePosition = 7;
			ultraGridColumn34.Width = 68;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance99.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance99;
			ultraGridColumn35.DataType = typeof(System.Decimal);
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.Caption = "%";
			ultraGridColumn35.Header.VisiblePosition = 9;
			ultraGridColumn35.Width = 76;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance100.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance100;
			ultraGridColumn36.DataType = typeof(System.Decimal);
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.Caption = "%";
			ultraGridColumn36.Header.VisiblePosition = 11;
			ultraGridColumn36.Width = 81;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance101.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance101;
			ultraGridColumn37.DataType = typeof(System.Decimal);
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.Caption = "%";
			ultraGridColumn37.Header.VisiblePosition = 13;
			ultraGridColumn37.Width = 67;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn37});
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance102.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance102;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance103.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance103;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance104.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance104;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance105.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance105;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance106.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance106;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance107.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance107;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance108.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance108;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance109.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance109;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance110;
			summarySettings22.DisplayFormat = "{0: #,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance111.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance111;
			summarySettings23.DisplayFormat = "{0: #,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance112.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance112;
			summarySettings24.DisplayFormat = "{0: #,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance113.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance113;
			summarySettings25.DisplayFormat = "{0: #,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance114.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance114;
			summarySettings26.DisplayFormat = "{0: #,##0.00}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings14,
																																															summarySettings15,
																																															summarySettings16,
																																															summarySettings17,
																																															summarySettings18,
																																															summarySettings19,
																																															summarySettings20,
																																															summarySettings21,
																																															summarySettings22,
																																															summarySettings23,
																																															summarySettings24,
																																															summarySettings25,
																																															summarySettings26});
			ultraGridBand3.SummaryFooterCaption = "TOTALES";
			this.uGridVentasMargen.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance115.ForeColor = System.Drawing.Color.Black;
			appearance115.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasMargen.DisplayLayout.Override.ActiveRowAppearance = appearance115;
			this.uGridVentasMargen.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVentasMargen.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVentasMargen.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance116.BackColor = System.Drawing.Color.Transparent;
			this.uGridVentasMargen.DisplayLayout.Override.CardAreaAppearance = appearance116;
			appearance117.ForeColor = System.Drawing.Color.Black;
			appearance117.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVentasMargen.DisplayLayout.Override.CellAppearance = appearance117;
			this.uGridVentasMargen.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance118.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance118.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance118.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance118.FontData.BoldAsString = "True";
			appearance118.FontData.Name = "Arial";
			appearance118.FontData.SizeInPoints = 8F;
			appearance118.ForeColor = System.Drawing.Color.White;
			appearance118.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVentasMargen.DisplayLayout.Override.HeaderAppearance = appearance118;
			appearance119.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance119.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance119.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasMargen.DisplayLayout.Override.RowAlternateAppearance = appearance119;
			appearance120.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance120.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance120.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasMargen.DisplayLayout.Override.RowSelectorAppearance = appearance120;
			appearance121.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance121.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance121.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVentasMargen.DisplayLayout.Override.SelectedRowAppearance = appearance121;
			this.uGridVentasMargen.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVentasMargen.Location = new System.Drawing.Point(8, 136);
			this.uGridVentasMargen.Name = "uGridVentasMargen";
			this.uGridVentasMargen.Size = new System.Drawing.Size(1226, 104);
			this.uGridVentasMargen.TabIndex = 809;
			this.uGridVentasMargen.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVentasMargen_InitializeLayout);
			// 
			// uGridOtrosIngresos
			// 
			this.uGridOtrosIngresos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridOtrosIngresos.DataSource = this.ultraDataSource5;
			appearance122.BackColor = System.Drawing.Color.White;
			this.uGridOtrosIngresos.DisplayLayout.Appearance = appearance122;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Header.Caption = "Año";
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 137;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn39.Header.VisiblePosition = 16;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn40.Header.Caption = "Otros Ingresos";
			ultraGridColumn40.Header.VisiblePosition = 1;
			ultraGridColumn40.Width = 180;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance123.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance123;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.Caption = "%";
			ultraGridColumn41.Header.VisiblePosition = 3;
			ultraGridColumn41.Width = 40;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance124.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance124;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.VisiblePosition = 4;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 80;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance125.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance125;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.VisiblePosition = 2;
			ultraGridColumn43.Width = 110;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance126.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance126;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 5;
			ultraGridColumn44.Width = 110;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance127.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance127;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.VisiblePosition = 7;
			ultraGridColumn45.Width = 110;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance128.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance128;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 9;
			ultraGridColumn46.Width = 110;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance129.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance129;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.VisiblePosition = 11;
			ultraGridColumn47.Width = 110;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance130.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance130;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.VisiblePosition = 13;
			ultraGridColumn48.Width = 110;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance131.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance131;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.VisiblePosition = 15;
			ultraGridColumn49.Width = 95;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance132.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance132;
			ultraGridColumn50.DataType = typeof(System.Decimal);
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.Caption = "%";
			ultraGridColumn50.Header.VisiblePosition = 6;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance133.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance133;
			ultraGridColumn51.DataType = typeof(System.Decimal);
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.Caption = "%";
			ultraGridColumn51.Header.VisiblePosition = 8;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance134.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance134;
			ultraGridColumn52.DataType = typeof(System.Decimal);
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.Caption = "%";
			ultraGridColumn52.Header.VisiblePosition = 10;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance135.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance135;
			ultraGridColumn53.DataType = typeof(System.Decimal);
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.Caption = "%";
			ultraGridColumn53.Header.VisiblePosition = 12;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance136.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance136;
			ultraGridColumn54.DataType = typeof(System.Decimal);
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.Caption = "%";
			ultraGridColumn54.Header.VisiblePosition = 14;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn54});
			ultraGridBand4.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance137.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings27.Appearance = appearance137;
			summarySettings27.DisplayFormat = "{0: #,##0.00}";
			summarySettings27.Hidden = true;
			summarySettings27.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance138.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings28.Appearance = appearance138;
			summarySettings28.DisplayFormat = "{0: #,##0.00}";
			summarySettings28.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance139.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings29.Appearance = appearance139;
			summarySettings29.DisplayFormat = "{0: #,##0.00}";
			summarySettings29.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance140.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings30.Appearance = appearance140;
			summarySettings30.DisplayFormat = "{0: #,##0.00}";
			summarySettings30.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance141.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings31.Appearance = appearance141;
			summarySettings31.DisplayFormat = "{0: #,##0.00}";
			summarySettings31.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance142.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings32.Appearance = appearance142;
			summarySettings32.DisplayFormat = "{0: #,##0.00}";
			summarySettings32.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance143.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings33.Appearance = appearance143;
			summarySettings33.DisplayFormat = "{0: #,##0.00}";
			summarySettings33.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance144.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings34.Appearance = appearance144;
			summarySettings34.DisplayFormat = "{0: #,##0.00}";
			summarySettings34.Hidden = true;
			summarySettings34.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance145.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings35.Appearance = appearance145;
			summarySettings35.DisplayFormat = "{0: #,##0.00}";
			summarySettings35.Hidden = true;
			summarySettings35.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance146.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings36.Appearance = appearance146;
			summarySettings36.DisplayFormat = "{0: #,##0.00}";
			summarySettings36.Hidden = true;
			summarySettings36.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance147.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings37.Appearance = appearance147;
			summarySettings37.DisplayFormat = "{0: #,##0.00}";
			summarySettings37.Hidden = true;
			summarySettings37.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance148.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings38.Appearance = appearance148;
			summarySettings38.DisplayFormat = "{0: #,##0.00}";
			summarySettings38.Hidden = true;
			summarySettings38.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance149.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings39.Appearance = appearance149;
			summarySettings39.DisplayFormat = "{0: #,##0.00}";
			summarySettings39.Hidden = true;
			summarySettings39.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings27,
																																															summarySettings28,
																																															summarySettings29,
																																															summarySettings30,
																																															summarySettings31,
																																															summarySettings32,
																																															summarySettings33,
																																															summarySettings34,
																																															summarySettings35,
																																															summarySettings36,
																																															summarySettings37,
																																															summarySettings38,
																																															summarySettings39});
			ultraGridBand4.SummaryFooterCaption = "TOTALES";
			this.uGridOtrosIngresos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance150.ForeColor = System.Drawing.Color.Black;
			appearance150.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOtrosIngresos.DisplayLayout.Override.ActiveRowAppearance = appearance150;
			this.uGridOtrosIngresos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridOtrosIngresos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridOtrosIngresos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance151.BackColor = System.Drawing.Color.Transparent;
			this.uGridOtrosIngresos.DisplayLayout.Override.CardAreaAppearance = appearance151;
			appearance152.ForeColor = System.Drawing.Color.Black;
			appearance152.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOtrosIngresos.DisplayLayout.Override.CellAppearance = appearance152;
			this.uGridOtrosIngresos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance153.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance153.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance153.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance153.FontData.BoldAsString = "True";
			appearance153.FontData.Name = "Arial";
			appearance153.FontData.SizeInPoints = 8F;
			appearance153.ForeColor = System.Drawing.Color.White;
			appearance153.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridOtrosIngresos.DisplayLayout.Override.HeaderAppearance = appearance153;
			appearance154.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance154.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance154.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOtrosIngresos.DisplayLayout.Override.RowAlternateAppearance = appearance154;
			appearance155.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance155.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance155.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOtrosIngresos.DisplayLayout.Override.RowSelectorAppearance = appearance155;
			appearance156.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance156.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance156.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOtrosIngresos.DisplayLayout.Override.SelectedRowAppearance = appearance156;
			this.uGridOtrosIngresos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridOtrosIngresos.Location = new System.Drawing.Point(8, 296);
			this.uGridOtrosIngresos.Name = "uGridOtrosIngresos";
			this.uGridOtrosIngresos.Size = new System.Drawing.Size(1226, 72);
			this.uGridOtrosIngresos.TabIndex = 812;
			this.uGridOtrosIngresos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridOtrosIngresos_KeyDown);
			this.uGridOtrosIngresos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridOtrosIngresos_InitializeLayout);
			// 
			// uGridGastos
			// 
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource6;
			appearance157.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance157;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance158.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance158;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance159.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance159;
			ultraGridColumn56.Header.VisiblePosition = 1;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn57.Header.VisiblePosition = 2;
			ultraGridColumn57.Width = 120;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance160.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance160;
			ultraGridColumn58.Header.VisiblePosition = 3;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 110;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance161.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance161;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.VisiblePosition = 4;
			ultraGridColumn59.Width = 110;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance162.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance162;
			ultraGridColumn60.Format = "#,##0.00";
			ultraGridColumn60.Header.VisiblePosition = 6;
			ultraGridColumn60.Width = 110;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance163.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance163;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.VisiblePosition = 8;
			ultraGridColumn61.Width = 110;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance164.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance164;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.VisiblePosition = 10;
			ultraGridColumn62.Width = 110;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance165.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance165;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.VisiblePosition = 12;
			ultraGridColumn63.Width = 110;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance166.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance166;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.VisiblePosition = 14;
			ultraGridColumn64.Width = 110;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance167.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance167;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.VisiblePosition = 16;
			ultraGridColumn65.Width = 110;
			appearance168.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance168;
			ultraGridColumn66.Header.VisiblePosition = 5;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance169.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance169;
			ultraGridColumn67.DataType = typeof(System.Decimal);
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.Caption = "%";
			ultraGridColumn67.Header.VisiblePosition = 7;
			ultraGridColumn68.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance170.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance170;
			ultraGridColumn68.DataType = typeof(System.Decimal);
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.Caption = "%";
			ultraGridColumn68.Header.VisiblePosition = 9;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance171.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance171;
			ultraGridColumn69.DataType = typeof(System.Decimal);
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.Caption = "%";
			ultraGridColumn69.Header.VisiblePosition = 11;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance172.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance172;
			ultraGridColumn70.DataType = typeof(System.Decimal);
			ultraGridColumn70.Format = "#,##0.00";
			ultraGridColumn70.Header.Caption = "%";
			ultraGridColumn70.Header.VisiblePosition = 13;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance173.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance173;
			ultraGridColumn71.DataType = typeof(System.Decimal);
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.Caption = "%";
			ultraGridColumn71.Header.VisiblePosition = 15;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn71});
			ultraGridBand5.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance174.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings40.Appearance = appearance174;
			summarySettings40.DisplayFormat = "{0: #,##0.00}";
			summarySettings40.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance175.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings41.Appearance = appearance175;
			summarySettings41.DisplayFormat = "{0: #,##0.00}";
			summarySettings41.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance176.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings42.Appearance = appearance176;
			summarySettings42.DisplayFormat = "{0: #,##0.00}";
			summarySettings42.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance177.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings43.Appearance = appearance177;
			summarySettings43.DisplayFormat = "{0: #,##0.00}";
			summarySettings43.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance178.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings44.Appearance = appearance178;
			summarySettings44.DisplayFormat = "{0: #,##0.00}";
			summarySettings44.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance179.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings45.Appearance = appearance179;
			summarySettings45.DisplayFormat = "{0: #,##0.00}";
			summarySettings45.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance180.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings46.Appearance = appearance180;
			summarySettings46.DisplayFormat = "{0: #,##0.00}";
			summarySettings46.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance181.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings47.Appearance = appearance181;
			summarySettings47.DisplayFormat = "{0: #,##0.00}";
			summarySettings47.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance182.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings48.Appearance = appearance182;
			summarySettings48.DisplayFormat = "{0: #,##0.00}";
			summarySettings48.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance183.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings49.Appearance = appearance183;
			summarySettings49.DisplayFormat = "{0: #,##0.00}";
			summarySettings49.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance184.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings50.Appearance = appearance184;
			summarySettings50.DisplayFormat = "{0: #,##0.00}";
			summarySettings50.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance185.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings51.Appearance = appearance185;
			summarySettings51.DisplayFormat = "{0: #,##0.00}";
			summarySettings51.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance186.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings52.Appearance = appearance186;
			summarySettings52.DisplayFormat = "{0: #,##0.00}";
			summarySettings52.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings40,
																																															summarySettings41,
																																															summarySettings42,
																																															summarySettings43,
																																															summarySettings44,
																																															summarySettings45,
																																															summarySettings46,
																																															summarySettings47,
																																															summarySettings48,
																																															summarySettings49,
																																															summarySettings50,
																																															summarySettings51,
																																															summarySettings52});
			ultraGridBand5.SummaryFooterCaption = "TOTALES";
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance187.ForeColor = System.Drawing.Color.Black;
			appearance187.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance187;
			this.uGridGastos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGastos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance188.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance188;
			appearance189.ForeColor = System.Drawing.Color.Black;
			appearance189.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.CellAppearance = appearance189;
			this.uGridGastos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance190.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance190.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance190.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance190.FontData.BoldAsString = "True";
			appearance190.FontData.Name = "Arial";
			appearance190.FontData.SizeInPoints = 8F;
			appearance190.ForeColor = System.Drawing.Color.White;
			appearance190.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance190;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance191.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance191.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance191.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance191;
			appearance192.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance192.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance192.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance192;
			appearance193.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance193.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance193.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance193;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 376);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(1215, 88);
			this.uGridGastos.TabIndex = 1272;
			this.uGridGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridGastos_KeyDown);
			this.uGridGastos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_ClickCellButton);
			this.uGridGastos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridGastos_InitializeLayout);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn75.DataType = typeof(System.Decimal);
			ultraDataColumn76.DataType = typeof(System.Decimal);
			ultraDataColumn77.DataType = typeof(System.Decimal);
			ultraDataColumn78.DataType = typeof(System.Decimal);
			ultraDataColumn79.DataType = typeof(System.Decimal);
			ultraDataColumn80.DataType = typeof(System.Decimal);
			ultraDataColumn81.DataType = typeof(System.Decimal);
			ultraDataColumn82.DataType = typeof(System.Decimal);
			ultraDataColumn83.DataType = typeof(System.Decimal);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn72,
																																 ultraDataColumn73,
																																 ultraDataColumn74,
																																 ultraDataColumn75,
																																 ultraDataColumn76,
																																 ultraDataColumn77,
																																 ultraDataColumn78,
																																 ultraDataColumn79,
																																 ultraDataColumn80,
																																 ultraDataColumn81,
																																 ultraDataColumn82,
																																 ultraDataColumn83});
			// 
			// label18
			// 
			this.label18.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.ForeColor = System.Drawing.Color.Firebrick;
			this.label18.Location = new System.Drawing.Point(24, 504);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(152, 16);
			this.label18.TabIndex = 1273;
			this.label18.Text = "GASTOS DE PATRIMONIO";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtGPEnero
			// 
			this.txtGPEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance194.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPEnero.Appearance = appearance194;
			this.txtGPEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPEnero.FormatString = "#,##0.00";
			this.txtGPEnero.Location = new System.Drawing.Point(96, 528);
			this.txtGPEnero.Name = "txtGPEnero";
			this.txtGPEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPEnero.PromptChar = ' ';
			this.txtGPEnero.Size = new System.Drawing.Size(80, 21);
			this.txtGPEnero.TabIndex = 1274;
			this.txtGPEnero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGPEnero_KeyDown);
			this.txtGPEnero.Click += new System.EventHandler(this.txtGPEnero_Click);
			this.txtGPEnero.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPEnero_Validating);
			this.txtGPEnero.ValueChanged += new System.EventHandler(this.txtGPEnero_ValueChanged);
			// 
			// txtGPFebrero
			// 
			this.txtGPFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance195.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPFebrero.Appearance = appearance195;
			this.txtGPFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPFebrero.FormatString = "#,##0.00";
			this.txtGPFebrero.Location = new System.Drawing.Point(184, 528);
			this.txtGPFebrero.Name = "txtGPFebrero";
			this.txtGPFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPFebrero.PromptChar = ' ';
			this.txtGPFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtGPFebrero.TabIndex = 1275;
			this.txtGPFebrero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGPFebrero_KeyDown);
			this.txtGPFebrero.Click += new System.EventHandler(this.txtGPFebrero_Click);
			this.txtGPFebrero.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPFebrero_Validating);
			// 
			// txtGPAbril
			// 
			this.txtGPAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance196.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPAbril.Appearance = appearance196;
			this.txtGPAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPAbril.FormatString = "#,##0.00";
			this.txtGPAbril.Location = new System.Drawing.Point(360, 528);
			this.txtGPAbril.Name = "txtGPAbril";
			this.txtGPAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPAbril.PromptChar = ' ';
			this.txtGPAbril.Size = new System.Drawing.Size(80, 21);
			this.txtGPAbril.TabIndex = 1277;
			this.txtGPAbril.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGPAbril_KeyDown);
			this.txtGPAbril.Click += new System.EventHandler(this.txtGPAbril_Click);
			this.txtGPAbril.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPAbril_Validating);
			// 
			// txtGPMarzo
			// 
			this.txtGPMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance197.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPMarzo.Appearance = appearance197;
			this.txtGPMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPMarzo.FormatString = "#,##0.00";
			this.txtGPMarzo.Location = new System.Drawing.Point(272, 528);
			this.txtGPMarzo.Name = "txtGPMarzo";
			this.txtGPMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPMarzo.PromptChar = ' ';
			this.txtGPMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtGPMarzo.TabIndex = 1276;
			this.txtGPMarzo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGPMarzo_KeyDown);
			this.txtGPMarzo.Click += new System.EventHandler(this.txtGPMarzo_Click);
			this.txtGPMarzo.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPMarzo_Validating);
			// 
			// txtGPAgosto
			// 
			this.txtGPAgosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance198.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPAgosto.Appearance = appearance198;
			this.txtGPAgosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPAgosto.FormatString = "#,##0.00";
			this.txtGPAgosto.Location = new System.Drawing.Point(712, 528);
			this.txtGPAgosto.Name = "txtGPAgosto";
			this.txtGPAgosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPAgosto.PromptChar = ' ';
			this.txtGPAgosto.Size = new System.Drawing.Size(80, 21);
			this.txtGPAgosto.TabIndex = 1281;
			this.txtGPAgosto.Visible = false;
			this.txtGPAgosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGPAgosto_KeyDown);
			this.txtGPAgosto.Click += new System.EventHandler(this.txtGPAgosto_Click);
			this.txtGPAgosto.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPAgosto_Validating);
			// 
			// txtGPJulio
			// 
			this.txtGPJulio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance199.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPJulio.Appearance = appearance199;
			this.txtGPJulio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPJulio.FormatString = "#,##0.00";
			this.txtGPJulio.Location = new System.Drawing.Point(624, 528);
			this.txtGPJulio.Name = "txtGPJulio";
			this.txtGPJulio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPJulio.PromptChar = ' ';
			this.txtGPJulio.Size = new System.Drawing.Size(80, 21);
			this.txtGPJulio.TabIndex = 1280;
			this.txtGPJulio.Visible = false;
			this.txtGPJulio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGPJulio_KeyDown);
			this.txtGPJulio.Click += new System.EventHandler(this.txtGPJulio_Click);
			this.txtGPJulio.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPJulio_Validating);
			// 
			// txtGPJunio
			// 
			this.txtGPJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance200.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPJunio.Appearance = appearance200;
			this.txtGPJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPJunio.FormatString = "#,##0.00";
			this.txtGPJunio.Location = new System.Drawing.Point(536, 528);
			this.txtGPJunio.Name = "txtGPJunio";
			this.txtGPJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPJunio.PromptChar = ' ';
			this.txtGPJunio.Size = new System.Drawing.Size(80, 21);
			this.txtGPJunio.TabIndex = 1279;
			this.txtGPJunio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGPJunio_KeyDown);
			this.txtGPJunio.Click += new System.EventHandler(this.txtGPJunio_Click);
			this.txtGPJunio.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPJunio_Validating);
			// 
			// txtGPMayo
			// 
			this.txtGPMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance201.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPMayo.Appearance = appearance201;
			this.txtGPMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPMayo.FormatString = "#,##0.00";
			this.txtGPMayo.Location = new System.Drawing.Point(448, 528);
			this.txtGPMayo.Name = "txtGPMayo";
			this.txtGPMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPMayo.PromptChar = ' ';
			this.txtGPMayo.Size = new System.Drawing.Size(80, 21);
			this.txtGPMayo.TabIndex = 1278;
			this.txtGPMayo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGPMayo_KeyDown);
			this.txtGPMayo.Click += new System.EventHandler(this.txtGPMayo_Click);
			this.txtGPMayo.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPMayo_Validating);
			// 
			// txtGPDiciembre
			// 
			this.txtGPDiciembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance202.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPDiciembre.Appearance = appearance202;
			this.txtGPDiciembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPDiciembre.FormatString = "#,##0.00";
			this.txtGPDiciembre.Location = new System.Drawing.Point(1064, 528);
			this.txtGPDiciembre.Name = "txtGPDiciembre";
			this.txtGPDiciembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPDiciembre.PromptChar = ' ';
			this.txtGPDiciembre.Size = new System.Drawing.Size(80, 21);
			this.txtGPDiciembre.TabIndex = 1285;
			this.txtGPDiciembre.Visible = false;
			this.txtGPDiciembre.Click += new System.EventHandler(this.txtGPDiciembre_Click);
			this.txtGPDiciembre.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPDiciembre_Validating);
			// 
			// txtGPNoviembre
			// 
			this.txtGPNoviembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance203.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPNoviembre.Appearance = appearance203;
			this.txtGPNoviembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPNoviembre.FormatString = "#,##0.00";
			this.txtGPNoviembre.Location = new System.Drawing.Point(976, 528);
			this.txtGPNoviembre.Name = "txtGPNoviembre";
			this.txtGPNoviembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPNoviembre.PromptChar = ' ';
			this.txtGPNoviembre.Size = new System.Drawing.Size(80, 21);
			this.txtGPNoviembre.TabIndex = 1284;
			this.txtGPNoviembre.Visible = false;
			this.txtGPNoviembre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGPNoviembre_KeyDown);
			this.txtGPNoviembre.Click += new System.EventHandler(this.txtGPNoviembre_Click);
			this.txtGPNoviembre.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPNoviembre_Validating);
			this.txtGPNoviembre.ValueChanged += new System.EventHandler(this.txtGPNoviembre_ValueChanged);
			// 
			// txtGPOctubre
			// 
			this.txtGPOctubre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance204.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPOctubre.Appearance = appearance204;
			this.txtGPOctubre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPOctubre.FormatString = "#,##0.00";
			this.txtGPOctubre.Location = new System.Drawing.Point(888, 528);
			this.txtGPOctubre.Name = "txtGPOctubre";
			this.txtGPOctubre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPOctubre.PromptChar = ' ';
			this.txtGPOctubre.Size = new System.Drawing.Size(80, 21);
			this.txtGPOctubre.TabIndex = 1283;
			this.txtGPOctubre.Visible = false;
			this.txtGPOctubre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGPOctubre_KeyDown);
			this.txtGPOctubre.Click += new System.EventHandler(this.txtGPOctubre_Click);
			this.txtGPOctubre.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPOctubre_Validating);
			// 
			// txtGPSeptiembre
			// 
			this.txtGPSeptiembre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance205.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGPSeptiembre.Appearance = appearance205;
			this.txtGPSeptiembre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGPSeptiembre.FormatString = "#,##0.00";
			this.txtGPSeptiembre.Location = new System.Drawing.Point(800, 528);
			this.txtGPSeptiembre.Name = "txtGPSeptiembre";
			this.txtGPSeptiembre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGPSeptiembre.PromptChar = ' ';
			this.txtGPSeptiembre.Size = new System.Drawing.Size(80, 21);
			this.txtGPSeptiembre.TabIndex = 1282;
			this.txtGPSeptiembre.Visible = false;
			this.txtGPSeptiembre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGPSeptiembre_KeyDown);
			this.txtGPSeptiembre.Click += new System.EventHandler(this.txtGPSeptiembre_Click);
			this.txtGPSeptiembre.Validating += new System.ComponentModel.CancelEventHandler(this.txtGPSeptiembre_Validating);
			// 
			// c1DataSet1
			// 
			this.c1DataSet1.BindingContextCtrl = this;
			this.c1DataSet1.DataLibrary = "LibFacturacion";
			this.c1DataSet1.DataLibraryUrl = "";
			this.c1DataSet1.DataSetDef = "dsSeteoF";
			this.c1DataSet1.Locale = new System.Globalization.CultureInfo("es-EC");
			this.c1DataSet1.Name = "c1DataSet1";
			this.c1DataSet1.SchemaClassName = "LibFacturacion.DataClass";
			this.c1DataSet1.SchemaDef = null;
			// 
			// ultraNumericEditor1
			// 
			this.ultraNumericEditor1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance206.ForeColor = System.Drawing.Color.Firebrick;
			appearance206.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.ultraNumericEditor1.Appearance = appearance206;
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.Enabled = false;
			this.ultraNumericEditor1.FormatString = "#,##0.00";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(976, 664);
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(80, 21);
			this.ultraNumericEditor1.TabIndex = 1297;
			this.ultraNumericEditor1.Visible = false;
			// 
			// ultraNumericEditor2
			// 
			this.ultraNumericEditor2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance207.ForeColor = System.Drawing.Color.Firebrick;
			appearance207.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.ultraNumericEditor2.Appearance = appearance207;
			this.ultraNumericEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor2.Enabled = false;
			this.ultraNumericEditor2.FormatString = "#,##0.00";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(1064, 664);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(80, 21);
			this.ultraNumericEditor2.TabIndex = 1296;
			this.ultraNumericEditor2.Visible = false;
			// 
			// ultraNumericEditor3
			// 
			this.ultraNumericEditor3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance208.ForeColor = System.Drawing.Color.Firebrick;
			appearance208.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.ultraNumericEditor3.Appearance = appearance208;
			this.ultraNumericEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor3.Enabled = false;
			this.ultraNumericEditor3.FormatString = "#,##0.00";
			this.ultraNumericEditor3.Location = new System.Drawing.Point(800, 664);
			this.ultraNumericEditor3.Name = "ultraNumericEditor3";
			this.ultraNumericEditor3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor3.PromptChar = ' ';
			this.ultraNumericEditor3.Size = new System.Drawing.Size(80, 21);
			this.ultraNumericEditor3.TabIndex = 1295;
			this.ultraNumericEditor3.Visible = false;
			// 
			// ultraNumericEditor4
			// 
			this.ultraNumericEditor4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance209.ForeColor = System.Drawing.Color.Firebrick;
			appearance209.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.ultraNumericEditor4.Appearance = appearance209;
			this.ultraNumericEditor4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor4.Enabled = false;
			this.ultraNumericEditor4.FormatString = "#,##0.00";
			this.ultraNumericEditor4.Location = new System.Drawing.Point(888, 664);
			this.ultraNumericEditor4.Name = "ultraNumericEditor4";
			this.ultraNumericEditor4.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor4.PromptChar = ' ';
			this.ultraNumericEditor4.Size = new System.Drawing.Size(80, 21);
			this.ultraNumericEditor4.TabIndex = 1294;
			this.ultraNumericEditor4.Visible = false;
			// 
			// ultraNumericEditor5
			// 
			this.ultraNumericEditor5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance210.ForeColor = System.Drawing.Color.Firebrick;
			appearance210.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.ultraNumericEditor5.Appearance = appearance210;
			this.ultraNumericEditor5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor5.Enabled = false;
			this.ultraNumericEditor5.FormatString = "#,##0.00";
			this.ultraNumericEditor5.Location = new System.Drawing.Point(1152, 664);
			this.ultraNumericEditor5.Name = "ultraNumericEditor5";
			this.ultraNumericEditor5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor5.PromptChar = ' ';
			this.ultraNumericEditor5.Size = new System.Drawing.Size(80, 21);
			this.ultraNumericEditor5.TabIndex = 1293;
			this.ultraNumericEditor5.Visible = false;
			// 
			// ultraNumericEditor6
			// 
			this.ultraNumericEditor6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance211.ForeColor = System.Drawing.Color.Firebrick;
			appearance211.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.ultraNumericEditor6.Appearance = appearance211;
			this.ultraNumericEditor6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor6.Enabled = false;
			this.ultraNumericEditor6.FormatString = "#,##0.00";
			this.ultraNumericEditor6.Location = new System.Drawing.Point(712, 664);
			this.ultraNumericEditor6.Name = "ultraNumericEditor6";
			this.ultraNumericEditor6.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor6.PromptChar = ' ';
			this.ultraNumericEditor6.Size = new System.Drawing.Size(80, 21);
			this.ultraNumericEditor6.TabIndex = 1292;
			this.ultraNumericEditor6.Visible = false;
			// 
			// txtGP22
			// 
			this.txtGP22.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance212.ForeColor = System.Drawing.Color.Firebrick;
			appearance212.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtGP22.Appearance = appearance212;
			this.txtGP22.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP22.Enabled = false;
			this.txtGP22.FormatString = "#,##0.00";
			this.txtGP22.Location = new System.Drawing.Point(448, 664);
			this.txtGP22.Name = "txtGP22";
			this.txtGP22.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP22.PromptChar = ' ';
			this.txtGP22.Size = new System.Drawing.Size(80, 21);
			this.txtGP22.TabIndex = 1291;
			// 
			// txtGP23
			// 
			this.txtGP23.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance213.ForeColor = System.Drawing.Color.Firebrick;
			appearance213.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtGP23.Appearance = appearance213;
			this.txtGP23.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP23.Enabled = false;
			this.txtGP23.FormatString = "#,##0.00";
			this.txtGP23.Location = new System.Drawing.Point(536, 664);
			this.txtGP23.Name = "txtGP23";
			this.txtGP23.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP23.PromptChar = ' ';
			this.txtGP23.Size = new System.Drawing.Size(80, 21);
			this.txtGP23.TabIndex = 1290;
			// 
			// txtGP20
			// 
			this.txtGP20.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance214.ForeColor = System.Drawing.Color.Firebrick;
			appearance214.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtGP20.Appearance = appearance214;
			this.txtGP20.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP20.Enabled = false;
			this.txtGP20.FormatString = "#,##0.00";
			this.txtGP20.Location = new System.Drawing.Point(272, 664);
			this.txtGP20.Name = "txtGP20";
			this.txtGP20.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP20.PromptChar = ' ';
			this.txtGP20.Size = new System.Drawing.Size(80, 21);
			this.txtGP20.TabIndex = 1289;
			// 
			// txtGP21
			// 
			this.txtGP21.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance215.ForeColor = System.Drawing.Color.Firebrick;
			appearance215.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtGP21.Appearance = appearance215;
			this.txtGP21.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP21.Enabled = false;
			this.txtGP21.FormatString = "#,##0.00";
			this.txtGP21.Location = new System.Drawing.Point(360, 664);
			this.txtGP21.Name = "txtGP21";
			this.txtGP21.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP21.PromptChar = ' ';
			this.txtGP21.Size = new System.Drawing.Size(80, 21);
			this.txtGP21.TabIndex = 1288;
			// 
			// txtGP18
			// 
			this.txtGP18.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance216.ForeColor = System.Drawing.Color.Firebrick;
			appearance216.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtGP18.Appearance = appearance216;
			this.txtGP18.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP18.Enabled = false;
			this.txtGP18.FormatString = "#,##0.00";
			this.txtGP18.Location = new System.Drawing.Point(96, 664);
			this.txtGP18.Name = "txtGP18";
			this.txtGP18.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP18.PromptChar = ' ';
			this.txtGP18.Size = new System.Drawing.Size(80, 21);
			this.txtGP18.TabIndex = 1287;
			// 
			// txtGP19
			// 
			this.txtGP19.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance217.ForeColor = System.Drawing.Color.Firebrick;
			appearance217.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtGP19.Appearance = appearance217;
			this.txtGP19.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP19.Enabled = false;
			this.txtGP19.FormatString = "#,##0.00";
			this.txtGP19.Location = new System.Drawing.Point(184, 664);
			this.txtGP19.Name = "txtGP19";
			this.txtGP19.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP19.PromptChar = ' ';
			this.txtGP19.Size = new System.Drawing.Size(80, 21);
			this.txtGP19.TabIndex = 1286;
			// 
			// txtUtilidadTotalP
			// 
			this.txtUtilidadTotalP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance218.FontData.BoldAsString = "True";
			appearance218.ForeColor = System.Drawing.Color.Firebrick;
			appearance218.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadTotalP.Appearance = appearance218;
			this.txtUtilidadTotalP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadTotalP.Enabled = false;
			this.txtUtilidadTotalP.FormatString = "#,##0.00";
			this.txtUtilidadTotalP.Location = new System.Drawing.Point(624, 664);
			this.txtUtilidadTotalP.Name = "txtUtilidadTotalP";
			this.txtUtilidadTotalP.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadTotalP.PromptChar = ' ';
			this.txtUtilidadTotalP.Size = new System.Drawing.Size(100, 21);
			this.txtUtilidadTotalP.TabIndex = 1298;
			// 
			// label19
			// 
			this.label19.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label19.AutoSize = true;
			this.label19.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label19.ForeColor = System.Drawing.Color.Firebrick;
			this.label19.Location = new System.Drawing.Point(56, 664);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(30, 16);
			this.label19.TabIndex = 1299;
			this.label19.Text = "(+) P";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label21
			// 
			this.label21.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label21.Location = new System.Drawing.Point(-8, 632);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(1240, 1);
			this.label21.TabIndex = 1301;
			// 
			// txtGP23T
			// 
			this.txtGP23T.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance219.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGP23T.Appearance = appearance219;
			this.txtGP23T.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP23T.Enabled = false;
			this.txtGP23T.FormatString = "#,##0.00";
			this.txtGP23T.Location = new System.Drawing.Point(536, 696);
			this.txtGP23T.Name = "txtGP23T";
			this.txtGP23T.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP23T.PromptChar = ' ';
			this.txtGP23T.Size = new System.Drawing.Size(80, 21);
			this.txtGP23T.TabIndex = 1307;
			// 
			// txtGP22T
			// 
			this.txtGP22T.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance220.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGP22T.Appearance = appearance220;
			this.txtGP22T.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP22T.Enabled = false;
			this.txtGP22T.FormatString = "#,##0.00";
			this.txtGP22T.Location = new System.Drawing.Point(448, 696);
			this.txtGP22T.Name = "txtGP22T";
			this.txtGP22T.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP22T.PromptChar = ' ';
			this.txtGP22T.Size = new System.Drawing.Size(80, 21);
			this.txtGP22T.TabIndex = 1306;
			// 
			// txtGP21T
			// 
			this.txtGP21T.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance221.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGP21T.Appearance = appearance221;
			this.txtGP21T.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP21T.Enabled = false;
			this.txtGP21T.FormatString = "#,##0.00";
			this.txtGP21T.Location = new System.Drawing.Point(360, 696);
			this.txtGP21T.Name = "txtGP21T";
			this.txtGP21T.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP21T.PromptChar = ' ';
			this.txtGP21T.Size = new System.Drawing.Size(80, 21);
			this.txtGP21T.TabIndex = 1305;
			// 
			// txtGP20T
			// 
			this.txtGP20T.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance222.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGP20T.Appearance = appearance222;
			this.txtGP20T.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP20T.Enabled = false;
			this.txtGP20T.FormatString = "#,##0.00";
			this.txtGP20T.Location = new System.Drawing.Point(272, 696);
			this.txtGP20T.Name = "txtGP20T";
			this.txtGP20T.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP20T.PromptChar = ' ';
			this.txtGP20T.Size = new System.Drawing.Size(80, 21);
			this.txtGP20T.TabIndex = 1304;
			// 
			// txtGP19T
			// 
			this.txtGP19T.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance223.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGP19T.Appearance = appearance223;
			this.txtGP19T.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP19T.Enabled = false;
			this.txtGP19T.FormatString = "#,##0.00";
			this.txtGP19T.Location = new System.Drawing.Point(184, 696);
			this.txtGP19T.Name = "txtGP19T";
			this.txtGP19T.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP19T.PromptChar = ' ';
			this.txtGP19T.Size = new System.Drawing.Size(80, 21);
			this.txtGP19T.TabIndex = 1303;
			// 
			// txtGP18T
			// 
			this.txtGP18T.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance224.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGP18T.Appearance = appearance224;
			this.txtGP18T.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGP18T.Enabled = false;
			this.txtGP18T.FormatString = "#,##0.00";
			this.txtGP18T.Location = new System.Drawing.Point(96, 696);
			this.txtGP18T.Name = "txtGP18T";
			this.txtGP18T.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGP18T.PromptChar = ' ';
			this.txtGP18T.Size = new System.Drawing.Size(80, 21);
			this.txtGP18T.TabIndex = 1302;
			// 
			// label20
			// 
			this.label20.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label20.AutoSize = true;
			this.label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label20.ForeColor = System.Drawing.Color.Firebrick;
			this.label20.Location = new System.Drawing.Point(56, 696);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(21, 16);
			this.label20.TabIndex = 1308;
			this.label20.Text = "UN";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtUtilidadTotalUN
			// 
			this.txtUtilidadTotalUN.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance225.FontData.BoldAsString = "True";
			appearance225.ForeColor = System.Drawing.Color.Firebrick;
			appearance225.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtUtilidadTotalUN.Appearance = appearance225;
			this.txtUtilidadTotalUN.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidadTotalUN.Enabled = false;
			this.txtUtilidadTotalUN.FormatString = "#,##0.00";
			this.txtUtilidadTotalUN.Location = new System.Drawing.Point(624, 696);
			this.txtUtilidadTotalUN.Name = "txtUtilidadTotalUN";
			this.txtUtilidadTotalUN.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidadTotalUN.PromptChar = ' ';
			this.txtUtilidadTotalUN.Size = new System.Drawing.Size(100, 21);
			this.txtUtilidadTotalUN.TabIndex = 1309;
			// 
			// lblprueba
			// 
			this.lblprueba.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblprueba.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblprueba.ForeColor = System.Drawing.Color.Firebrick;
			this.lblprueba.Location = new System.Drawing.Point(952, 504);
			this.lblprueba.Name = "lblprueba";
			this.lblprueba.Size = new System.Drawing.Size(152, 16);
			this.lblprueba.TabIndex = 1310;
			this.lblprueba.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblprueba.Visible = false;
			// 
			// ultraDataSource8
			// 
			ultraDataColumn85.DataType = typeof(System.Decimal);
			ultraDataColumn86.DataType = typeof(System.Decimal);
			ultraDataColumn87.DataType = typeof(System.Decimal);
			ultraDataColumn88.DataType = typeof(System.Decimal);
			ultraDataColumn89.DataType = typeof(System.Decimal);
			ultraDataColumn90.DataType = typeof(System.Decimal);
			ultraDataColumn91.DataType = typeof(System.Decimal);
			ultraDataColumn92.DataType = typeof(System.Decimal);
			ultraDataColumn93.DataType = typeof(System.Decimal);
			ultraDataColumn94.DataType = typeof(System.Decimal);
			ultraDataColumn95.DataType = typeof(System.Decimal);
			ultraDataColumn96.DataType = typeof(System.Decimal);
			ultraDataColumn97.DataType = typeof(System.Decimal);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn84,
																																 ultraDataColumn85,
																																 ultraDataColumn86,
																																 ultraDataColumn87,
																																 ultraDataColumn88,
																																 ultraDataColumn89,
																																 ultraDataColumn90,
																																 ultraDataColumn91,
																																 ultraDataColumn92,
																																 ultraDataColumn93,
																																 ultraDataColumn94,
																																 ultraDataColumn95,
																																 ultraDataColumn96,
																																 ultraDataColumn97});
			// 
			// UGridUb
			// 
			this.UGridUb.Cursor = System.Windows.Forms.Cursors.Default;
			this.UGridUb.DataSource = this.ultraDataSource8;
			appearance226.BackColor = System.Drawing.Color.White;
			this.UGridUb.DisplayLayout.Appearance = appearance226;
			ultraGridColumn72.Header.VisiblePosition = 0;
			ultraGridColumn73.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance227.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance227;
			ultraGridColumn73.Format = "#,##0.00";
			ultraGridColumn73.Header.VisiblePosition = 2;
			ultraGridColumn74.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance228.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance228;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.VisiblePosition = 1;
			ultraGridColumn74.Width = 110;
			ultraGridColumn75.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance229.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance229;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.VisiblePosition = 3;
			ultraGridColumn75.Width = 110;
			ultraGridColumn76.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance230.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance230;
			ultraGridColumn76.Format = "#,##0.00";
			ultraGridColumn76.Header.VisiblePosition = 5;
			ultraGridColumn77.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance231.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance231;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.VisiblePosition = 7;
			ultraGridColumn77.Width = 110;
			ultraGridColumn78.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance232.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance232;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.VisiblePosition = 9;
			ultraGridColumn78.Width = 110;
			ultraGridColumn79.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance233.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance233;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.VisiblePosition = 11;
			ultraGridColumn79.Width = 110;
			ultraGridColumn80.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance234.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance234;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.VisiblePosition = 13;
			ultraGridColumn80.Width = 110;
			ultraGridColumn81.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance235.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn81.CellAppearance = appearance235;
			ultraGridColumn81.Format = "#,##0.00";
			ultraGridColumn81.Header.Caption = "%";
			ultraGridColumn81.Header.VisiblePosition = 4;
			ultraGridColumn82.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance236.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn82.CellAppearance = appearance236;
			ultraGridColumn82.Format = "#,##0.00";
			ultraGridColumn82.Header.Caption = "%";
			ultraGridColumn82.Header.VisiblePosition = 6;
			ultraGridColumn83.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance237.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn83.CellAppearance = appearance237;
			ultraGridColumn83.Format = "#,##0.00";
			ultraGridColumn83.Header.Caption = "%";
			ultraGridColumn83.Header.VisiblePosition = 8;
			ultraGridColumn84.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance238.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn84.CellAppearance = appearance238;
			ultraGridColumn84.Format = "#,##0.00";
			ultraGridColumn84.Header.Caption = "%";
			ultraGridColumn84.Header.VisiblePosition = 10;
			ultraGridColumn85.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance239.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn85.CellAppearance = appearance239;
			ultraGridColumn85.Format = "#,##0.00";
			ultraGridColumn85.Header.Caption = "%";
			ultraGridColumn85.Header.VisiblePosition = 12;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn85});
			ultraGridBand6.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance240.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings53.Appearance = appearance240;
			summarySettings53.DisplayFormat = "{0: #,##0.00}";
			summarySettings53.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance241.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings54.Appearance = appearance241;
			summarySettings54.DisplayFormat = "{0: #,##0.00}";
			summarySettings54.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance242.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings55.Appearance = appearance242;
			summarySettings55.DisplayFormat = "{0: #,##0.00}";
			summarySettings55.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance243.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings56.Appearance = appearance243;
			summarySettings56.DisplayFormat = "{0: #,##0.00}";
			summarySettings56.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance244.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings57.Appearance = appearance244;
			summarySettings57.DisplayFormat = "{0: #,##0.00}";
			summarySettings57.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance245.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings58.Appearance = appearance245;
			summarySettings58.DisplayFormat = "{0: #,##0.00}";
			summarySettings58.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance246.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings59.Appearance = appearance246;
			summarySettings59.DisplayFormat = "{0: #,##0.00}";
			summarySettings59.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance247.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings60.Appearance = appearance247;
			summarySettings60.DisplayFormat = "{0: #,##0.00}";
			summarySettings60.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance248.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings61.Appearance = appearance248;
			summarySettings61.DisplayFormat = "{0: #,##0.00}";
			summarySettings61.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance249.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings62.Appearance = appearance249;
			summarySettings62.DisplayFormat = "{0: #,##0.00}";
			summarySettings62.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance250.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings63.Appearance = appearance250;
			summarySettings63.DisplayFormat = "{0: #,##0.00}";
			summarySettings63.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance251.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings64.Appearance = appearance251;
			summarySettings64.DisplayFormat = "{0: #,##0.00}";
			summarySettings64.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance252.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings65.Appearance = appearance252;
			summarySettings65.DisplayFormat = "{0: #,##0.00}";
			summarySettings65.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand6.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings53,
																																															summarySettings54,
																																															summarySettings55,
																																															summarySettings56,
																																															summarySettings57,
																																															summarySettings58,
																																															summarySettings59,
																																															summarySettings60,
																																															summarySettings61,
																																															summarySettings62,
																																															summarySettings63,
																																															summarySettings64,
																																															summarySettings65});
			ultraGridBand6.SummaryFooterCaption = "TOTALES";
			this.UGridUb.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance253.ForeColor = System.Drawing.Color.Black;
			appearance253.ForeColorDisabled = System.Drawing.Color.Black;
			this.UGridUb.DisplayLayout.Override.ActiveRowAppearance = appearance253;
			this.UGridUb.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.UGridUb.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.UGridUb.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance254.BackColor = System.Drawing.Color.Transparent;
			this.UGridUb.DisplayLayout.Override.CardAreaAppearance = appearance254;
			appearance255.ForeColor = System.Drawing.Color.Black;
			appearance255.ForeColorDisabled = System.Drawing.Color.Black;
			this.UGridUb.DisplayLayout.Override.CellAppearance = appearance255;
			this.UGridUb.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance256.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance256.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance256.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance256.FontData.BoldAsString = "True";
			appearance256.FontData.Name = "Arial";
			appearance256.FontData.SizeInPoints = 8F;
			appearance256.ForeColor = System.Drawing.Color.White;
			appearance256.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.UGridUb.DisplayLayout.Override.HeaderAppearance = appearance256;
			this.UGridUb.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance257.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance257.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance257.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UGridUb.DisplayLayout.Override.RowAlternateAppearance = appearance257;
			appearance258.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance258.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance258.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UGridUb.DisplayLayout.Override.RowSelectorAppearance = appearance258;
			appearance259.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance259.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance259.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.UGridUb.DisplayLayout.Override.SelectedRowAppearance = appearance259;
			this.UGridUb.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.UGridUb.Location = new System.Drawing.Point(736, 520);
			this.UGridUb.Name = "UGridUb";
			this.UGridUb.Size = new System.Drawing.Size(488, 200);
			this.UGridUb.TabIndex = 1311;
			this.UGridUb.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.UGridUb_InitializeLayout);
			// 
			// label22
			// 
			this.label22.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label22.ForeColor = System.Drawing.Color.Firebrick;
			this.label22.Location = new System.Drawing.Point(8, 472);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(88, 24);
			this.label22.TabIndex = 1312;
			this.label22.Text = "Perdida de Cartera";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtdETJunio
			// 
			this.txtdETJunio.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance260.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdETJunio.Appearance = appearance260;
			this.txtdETJunio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdETJunio.FormatString = "#,##0.00";
			this.txtdETJunio.Location = new System.Drawing.Point(544, 480);
			this.txtdETJunio.Name = "txtdETJunio";
			this.txtdETJunio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtdETJunio.PromptChar = ' ';
			this.txtdETJunio.Size = new System.Drawing.Size(80, 21);
			this.txtdETJunio.TabIndex = 1318;
			this.txtdETJunio.Validating += new System.ComponentModel.CancelEventHandler(this.txtdETJunio_Validating);
			// 
			// txtdETMayo
			// 
			this.txtdETMayo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance261.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdETMayo.Appearance = appearance261;
			this.txtdETMayo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdETMayo.FormatString = "#,##0.00";
			this.txtdETMayo.Location = new System.Drawing.Point(456, 480);
			this.txtdETMayo.Name = "txtdETMayo";
			this.txtdETMayo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtdETMayo.PromptChar = ' ';
			this.txtdETMayo.Size = new System.Drawing.Size(80, 21);
			this.txtdETMayo.TabIndex = 1317;
			this.txtdETMayo.Validating += new System.ComponentModel.CancelEventHandler(this.txtdETMayo_Validating);
			// 
			// txtdETAbril
			// 
			this.txtdETAbril.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance262.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdETAbril.Appearance = appearance262;
			this.txtdETAbril.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdETAbril.FormatString = "#,##0.00";
			this.txtdETAbril.Location = new System.Drawing.Point(368, 480);
			this.txtdETAbril.Name = "txtdETAbril";
			this.txtdETAbril.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtdETAbril.PromptChar = ' ';
			this.txtdETAbril.Size = new System.Drawing.Size(80, 21);
			this.txtdETAbril.TabIndex = 1316;
			this.txtdETAbril.Validating += new System.ComponentModel.CancelEventHandler(this.txtdETAbril_Validating);
			// 
			// txtdETMarzo
			// 
			this.txtdETMarzo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance263.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdETMarzo.Appearance = appearance263;
			this.txtdETMarzo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdETMarzo.FormatString = "#,##0.00";
			this.txtdETMarzo.Location = new System.Drawing.Point(280, 480);
			this.txtdETMarzo.Name = "txtdETMarzo";
			this.txtdETMarzo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtdETMarzo.PromptChar = ' ';
			this.txtdETMarzo.Size = new System.Drawing.Size(80, 21);
			this.txtdETMarzo.TabIndex = 1315;
			this.txtdETMarzo.Validating += new System.ComponentModel.CancelEventHandler(this.txtdETMarzo_Validating);
			// 
			// txtdETFebrero
			// 
			this.txtdETFebrero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance264.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdETFebrero.Appearance = appearance264;
			this.txtdETFebrero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdETFebrero.FormatString = "#,##0.00";
			this.txtdETFebrero.Location = new System.Drawing.Point(192, 480);
			this.txtdETFebrero.Name = "txtdETFebrero";
			this.txtdETFebrero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtdETFebrero.PromptChar = ' ';
			this.txtdETFebrero.Size = new System.Drawing.Size(80, 21);
			this.txtdETFebrero.TabIndex = 1314;
			this.txtdETFebrero.Validating += new System.ComponentModel.CancelEventHandler(this.txtdETFebrero_Validating);
			// 
			// txtdETEnero
			// 
			this.txtdETEnero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance265.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdETEnero.Appearance = appearance265;
			this.txtdETEnero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdETEnero.FormatString = "#,##0.00";
			this.txtdETEnero.Location = new System.Drawing.Point(104, 480);
			this.txtdETEnero.Name = "txtdETEnero";
			this.txtdETEnero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtdETEnero.PromptChar = ' ';
			this.txtdETEnero.Size = new System.Drawing.Size(80, 21);
			this.txtdETEnero.TabIndex = 1313;
			this.txtdETEnero.Validating += new System.ComponentModel.CancelEventHandler(this.txtdETEnero_Validating);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource8;
			appearance266.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance266;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn86.Header.Caption = "Porcentaje";
			ultraGridColumn86.Header.VisiblePosition = 0;
			ultraGridColumn86.Width = 204;
			ultraGridColumn87.Header.VisiblePosition = 1;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn87.Width = 82;
			ultraGridColumn88.Header.VisiblePosition = 2;
			ultraGridColumn88.Width = 179;
			ultraGridColumn89.Header.VisiblePosition = 3;
			ultraGridColumn89.Width = 159;
			ultraGridColumn90.Header.VisiblePosition = 4;
			ultraGridColumn90.Width = 159;
			ultraGridColumn91.Header.VisiblePosition = 5;
			ultraGridColumn91.Width = 159;
			ultraGridColumn92.Header.VisiblePosition = 6;
			ultraGridColumn92.Width = 159;
			ultraGridColumn93.Header.VisiblePosition = 7;
			ultraGridColumn93.Width = 159;
			ultraGridColumn94.Header.VisiblePosition = 8;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Width = 82;
			ultraGridColumn95.Header.VisiblePosition = 9;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn95.Width = 82;
			ultraGridColumn96.Header.VisiblePosition = 10;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn96.Width = 82;
			ultraGridColumn97.Header.VisiblePosition = 11;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 82;
			ultraGridColumn98.Header.VisiblePosition = 12;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn98.Width = 82;
			ultraGridColumn99.Header.VisiblePosition = 13;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn99.Width = 82;
			ultraGridBand7.Columns.AddRange(new object[] {
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
																										 ultraGridColumn99});
			ultraGridBand7.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand7.SummaryFooterCaption = "TOTALES";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance267.ForeColor = System.Drawing.Color.Black;
			appearance267.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance267;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance268.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance268;
			appearance269.ForeColor = System.Drawing.Color.Black;
			appearance269.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance269;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance270.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance270.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance270.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance270.FontData.BoldAsString = "True";
			appearance270.FontData.Name = "Arial";
			appearance270.FontData.SizeInPoints = 8F;
			appearance270.ForeColor = System.Drawing.Color.White;
			appearance270.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance270;
			appearance271.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance271.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance271.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance271;
			appearance272.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance272.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance272.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance272;
			appearance273.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance273.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance273.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance273;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 248);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1216, 40);
			this.ultraGrid1.TabIndex = 1319;
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// frmPresupuestoReal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1240, 726);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtdETJunio);
			this.Controls.Add(this.txtdETMayo);
			this.Controls.Add(this.txtdETAbril);
			this.Controls.Add(this.txtdETMarzo);
			this.Controls.Add(this.txtdETFebrero);
			this.Controls.Add(this.txtdETEnero);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.UGridUb);
			this.Controls.Add(this.lblprueba);
			this.Controls.Add(this.txtUtilidadTotalUN);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.txtGP23T);
			this.Controls.Add(this.txtGP22T);
			this.Controls.Add(this.txtGP21T);
			this.Controls.Add(this.txtGP20T);
			this.Controls.Add(this.txtGP19T);
			this.Controls.Add(this.txtGP18T);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.ultraNumericEditor2);
			this.Controls.Add(this.ultraNumericEditor3);
			this.Controls.Add(this.ultraNumericEditor4);
			this.Controls.Add(this.ultraNumericEditor5);
			this.Controls.Add(this.ultraNumericEditor6);
			this.Controls.Add(this.txtGP22);
			this.Controls.Add(this.txtGP23);
			this.Controls.Add(this.txtGP20);
			this.Controls.Add(this.txtGP21);
			this.Controls.Add(this.txtGP18);
			this.Controls.Add(this.txtGP19);
			this.Controls.Add(this.txtUtilidadTotalP);
			this.Controls.Add(this.txtGPDiciembre);
			this.Controls.Add(this.txtGPNoviembre);
			this.Controls.Add(this.txtGPOctubre);
			this.Controls.Add(this.txtGPSeptiembre);
			this.Controls.Add(this.txtGPAgosto);
			this.Controls.Add(this.txtGPJulio);
			this.Controls.Add(this.txtGPJunio);
			this.Controls.Add(this.txtGPMayo);
			this.Controls.Add(this.txtGPAbril);
			this.Controls.Add(this.txtGPMarzo);
			this.Controls.Add(this.txtGPFebrero);
			this.Controls.Add(this.txtGPEnero);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.uGridGastos);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtUtilidadNoviembre);
			this.Controls.Add(this.txtUtilidadDiciembre);
			this.Controls.Add(this.txtUtilidadSeptiembre);
			this.Controls.Add(this.txtUtilidadOctubre);
			this.Controls.Add(this.txtUtilidadJulio);
			this.Controls.Add(this.txtUtilidadAgosto);
			this.Controls.Add(this.label17);
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
			this.Controls.Add(this.txtUtilidadMayo);
			this.Controls.Add(this.txtUtilidadJunio);
			this.Controls.Add(this.txtUtilidadMarzo);
			this.Controls.Add(this.txtUtilidadAbril);
			this.Controls.Add(this.txtUtilidadEnero);
			this.Controls.Add(this.txtUtilidadFebrero);
			this.Controls.Add(this.txtGastosNoviembre);
			this.Controls.Add(this.txtGastosDiciembre);
			this.Controls.Add(this.txtIngresosDiciembre);
			this.Controls.Add(this.txtIngresosNoviembre);
			this.Controls.Add(this.txtGastosSeptiembre);
			this.Controls.Add(this.txtGastosOctubre);
			this.Controls.Add(this.txtIngresosOctubre);
			this.Controls.Add(this.txtIngresosSeptiembre);
			this.Controls.Add(this.txtGastosJulio);
			this.Controls.Add(this.txtGastosAgosto);
			this.Controls.Add(this.txtIngresosAgosto);
			this.Controls.Add(this.txtIngresosJulio);
			this.Controls.Add(this.txtGastosMayo);
			this.Controls.Add(this.txtGastosJunio);
			this.Controls.Add(this.txtIngresosJunio);
			this.Controls.Add(this.txtIngresosMayo);
			this.Controls.Add(this.txtGastosMarzo);
			this.Controls.Add(this.txtGastosAbril);
			this.Controls.Add(this.txtIngresosAbril);
			this.Controls.Add(this.txtIngresosMarzo);
			this.Controls.Add(this.txtGastosEnero);
			this.Controls.Add(this.txtGastosFebrero);
			this.Controls.Add(this.txtIngresosFebrero);
			this.Controls.Add(this.txtIngresosEnero);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.txtAño);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.uGridVentas);
			this.Controls.Add(this.uGridVentasPrecios);
			this.Controls.Add(this.uGridVentasMargen);
			this.Controls.Add(this.uGridOtrosIngresos);
			this.Controls.Add(this.btnProyeccion);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtUtilidadTotal);
			this.Controls.Add(this.txtGastosTotal);
			this.Controls.Add(this.txtIngresosTotal);
			this.Controls.Add(this.label2);
			this.KeyPreview = true;
			this.Name = "frmPresupuestoReal";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Presupuesto";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPresupuestoReal_KeyDown);
			this.Load += new System.EventHandler(this.frmPresupuestoReal_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresosEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasPrecios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVentasMargen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOtrosIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPAgosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPJulio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPDiciembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPNoviembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPOctubre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGPSeptiembre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataSet1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP22)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP23)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP21)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP18)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP19)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadTotalP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP23T)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP22T)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP21T)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP20T)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP19T)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGP18T)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidadTotalUN)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.UGridUb)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETJunio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETMayo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETAbril)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETMarzo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETFebrero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdETEnero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		DateTime dtFechaInicial = DateTime.Today;
		DateTime dtFechaFinal = DateTime.Today;

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			dtFechaInicial = DateTime.Parse("01/01/" + this.txtAño.Value);
			dtFechaFinal = DateTime.Parse("31/12/" + this.txtAño.Value);

			
			DateTime dtDesde = DateTime.Parse("01/01/" + this.txtAño.Value);
			DateTime dtHasta = dtDesde;

			if (dtFechaFinal >= dtHasta) dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			//			string sSQLra = string.Format("Exec [ProyeccionMensualVentasRealAnual] '{0}', '{1}'", dtFechaInicial.ToString("yyyyMMdd"), dtFechaFinal.ToString("yyyyMMdd"));
			//			this.uGridVentas.DataSource = FuncionesProcedimientos.dtGeneral(sSQLra);
			//	

			string sSQL = string.Format("Exec [ProyeccionMensualVentasRealPeriodo] '{0}', '{1}'", dtFechaInicial.ToString("yyyyMMdd"), dtFechaFinal.ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneralModulo (sSQL, this.uGridVentasPrecios);
			this.sum_fields(this.uGridVentasPrecios);

			string sSQLut = string.Format("Exec [PROYECCIONMENSUALMARGENREALPEriodo] '{0}', '{1}'", dtFechaInicial.ToString("yyyyMMdd"), dtFechaFinal.ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneralModulo (sSQLut, this.uGridVentasMargen);
			this.sum_fields(this.uGridVentasMargen);

			string sSQLutGo = string.Format("Exec [ProyeccionRealidadPresupuestoAnualGastosAlmacenesREalOtrosGastosPeriodo] '{0}', '{1}', {2}", dtFechaInicial.ToString("yyyyMMdd"), dtFechaFinal.ToString("yyyyMMdd"), this.txtAño.Value);
			FuncionesProcedimientos.dsGeneralModulo (sSQLutGo, this.uGridOtrosIngresos);	
			this.sum_fields(this.uGridOtrosIngresos);

			string sSQLutG = string.Format("Exec [ProyeccionRealidadPresupuestoAnualGastosAlmacenesREalPeriodo] '{0}', '{1}', {2}", dtFechaInicial.ToString("yyyyMMdd"), dtFechaFinal.ToString("yyyyMMdd"), this.txtAño.Value);
			FuncionesProcedimientos.dsGeneralModulo (sSQLutG, this.uGridGastos);
			this.sum_fields(this.uGridGastos);

			string sSQLUb = string.Format("Exec [ReporteUtilidadBrutaAnual] ");
			FuncionesProcedimientos.dsGeneralModulo (sSQLUb, this.UGridUb);

			string sSQLUbpo = string.Format("Exec [ReporteUtilidadBrutaAnualPorcentaje] ");
			FuncionesProcedimientos.dsGeneralModulo (sSQLUbpo, this.ultraGrid1);

			this.utilidaBruta ();
			this.sum_fields(this.UGridUb);

			this.Consulta(Convert.ToInt32(this.txtAño.Value));

			this.Resultado();
			this.ResultadoMargen();
			//this.DiseñoGridSumatoria(this.uGridGastos);



		}

		private void ResultadoMargen ()
		{
			decimal VPorcentaje2018 = 0.00m;
			decimal VPorcentaje2019 = 0.00m;
			decimal VPorcentaje2020 = 0.00m;
			decimal VPorcentaje2021 = 0.00m;
			decimal VPorcentaje2022 = 0.00m;
			decimal VPorcentaje2023 = 0.00m;

			decimal CtPorcentaje2018 = 0.00m;
			decimal CtPorcentaje2019 = 0.00m;
			decimal CtPorcentaje2020 = 0.00m;
			decimal CtPorcentaje2021 = 0.00m;
			decimal CtPorcentaje2022 = 0.00m;
			decimal CtPorcentaje2023 = 0.00m;

			decimal TtPorcentaje2018 = 0.00m;
			decimal TtPorcentaje2019 = 0.00m;
			decimal TtPorcentaje2020 = 0.00m;
			decimal TtPorcentaje2021 = 0.00m;
			decimal TtPorcentaje2022 = 0.00m;
			decimal TtPorcentaje2023 = 0.00m;


			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasPrecios.Rows.All)
			{
				VPorcentaje2018 =  VPorcentaje2018 + Convert.ToDecimal(dr.Cells["P_2018"].Value);
				VPorcentaje2019 =  VPorcentaje2019 + Convert.ToDecimal(dr.Cells["P_2019"].Value);
				VPorcentaje2020 =  VPorcentaje2020 + Convert.ToDecimal(dr.Cells["P_2020"].Value);
				VPorcentaje2021 =  VPorcentaje2021 + Convert.ToDecimal(dr.Cells["P_2021"].Value);
				VPorcentaje2022 =  VPorcentaje2022 + Convert.ToDecimal(dr.Cells["P_2022"].Value);
				VPorcentaje2023 =  VPorcentaje2023 + Convert.ToDecimal(dr.Cells["P_2023"].Value);
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasMargen.Rows.All)
			{
				CtPorcentaje2018 =  CtPorcentaje2018 + Convert.ToDecimal(dr.Cells["P_2018"].Value);
				CtPorcentaje2019 =  CtPorcentaje2019 + Convert.ToDecimal(dr.Cells["P_2019"].Value);
				CtPorcentaje2020 =  CtPorcentaje2020 + Convert.ToDecimal(dr.Cells["P_2020"].Value);
				CtPorcentaje2021 =  CtPorcentaje2021 + Convert.ToDecimal(dr.Cells["P_2021"].Value);
				CtPorcentaje2022 =  CtPorcentaje2022 + Convert.ToDecimal(dr.Cells["P_2022"].Value);
				CtPorcentaje2023 =  CtPorcentaje2023 + Convert.ToDecimal(dr.Cells["P_2023"].Value);
			}
		
			if(VPorcentaje2018 > 0)
			{
				TtPorcentaje2018 = CtPorcentaje2018/ VPorcentaje2018  ;
			}
			if(VPorcentaje2019 > 0)
			{
				TtPorcentaje2019 = CtPorcentaje2019 / VPorcentaje2019  ;
			}
			if(VPorcentaje2020 > 0)
			{

				TtPorcentaje2020 = CtPorcentaje2020 /VPorcentaje2020 ;
			}
			if(VPorcentaje2021 > 0)
			{
				TtPorcentaje2021 = CtPorcentaje2021 /VPorcentaje2021  ;
			}
			if(VPorcentaje2022 > 0)
			{
       
				TtPorcentaje2022 = CtPorcentaje2022 /VPorcentaje2022 ;
			}
			if(VPorcentaje2023 > 0)
			{
				TtPorcentaje2023 = CtPorcentaje2023 / VPorcentaje2023 ;
			}  

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				dr.Cells["P_2018"].Value =  Math.Round(TtPorcentaje2018,2);
				dr.Cells["P_2019"].Value =  Math.Round(TtPorcentaje2019,2);
				dr.Cells["P_2020"].Value =  Math.Round(TtPorcentaje2020,2);
				dr.Cells["P_2021"].Value =  Math.Round(TtPorcentaje2021,2);
				dr.Cells["P_2022"].Value =  Math.Round(TtPorcentaje2022,2);
				dr.Cells["P_2023"].Value =  Math.Round(TtPorcentaje2023,2);


			}
		}


		private void DiseñoGridSumatoria(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
		
			if (this.uGridGastos.Rows.Count > 0)
			{				
				int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;					
					
				string nombrCol = "";

				for (int j = 2; j < iContColumnas; j++)
				{
					nombrCol = this.uGridGastos.DisplayLayout.Bands[0].Columns[j].Header.Caption;	

					if( nombrCol == "P_2018" || nombrCol == "P_2019" || nombrCol == "P_2020" || nombrCol == "P_2021" || nombrCol == "P_2022" || nombrCol == "P_2023")
					{
						this.uGridGastos.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
						this.uGridGastos.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
						this.uGridGastos.DisplayLayout.Bands[0].Columns[nombrCol].Width = 130;
						//						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0.00";
						this.uGridGastos.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;	
					}
				}
			}

		}


		private void utilidaBruta ()
		{
		    
			decimal Ven_dP_2018 = 0.00m;
			decimal Ven_dP_2019 = 0.00m;
			decimal Ven_dP_2020 = 0.00m;
			decimal Ven_dP_2021 = 0.00m;
			decimal Ven_dP_2022 = 0.00m;
			decimal Ven_dP_2023 = 0.00m;

			decimal TVen_dP_2018 = 0.00m;
			decimal TVen_dP_2019 = 0.00m;
			decimal TVen_dP_2020 = 0.00m;
			decimal TVen_dP_2021 = 0.00m;
			decimal TVen_dP_2022 = 0.00m;
			decimal TVen_dP_2023 = 0.00m;

			decimal CVen_dP_2018 = 0.00m;
			decimal CVen_dP_2019 = 0.00m;
			decimal CVen_dP_2020 = 0.00m;
			decimal CVen_dP_2021 = 0.00m;
			decimal CVen_dP_2022 = 0.00m;
			decimal CVen_dP_2023 = 0.00m;

			decimal dVen_dP_2018 = 0.00m;
			decimal dVen_dP_2019 = 0.00m;
			decimal dVen_dP_2020 = 0.00m;
			decimal dVen_dP_2021 = 0.00m;
			decimal dVen_dP_2022 = 0.00m;
			decimal dVen_dP_2023 = 0.00m;

			string sColumna = "";
			int iContColumnas = (int)this.uGridVentasPrecios.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasPrecios.Rows.All)
			{
				
				if ( Convert.ToString( dr.Cells["Venta"].Value) == "EFECTIVO")
				{
					Ven_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2018"].Value);
					Ven_dP_2019 = Convert.ToDecimal ( dr.Cells["p_2019"].Value);
					Ven_dP_2020 = Convert.ToDecimal ( dr.Cells["p_2020"].Value);
					Ven_dP_2021 = Convert.ToDecimal ( dr.Cells["p_2021"].Value);
					Ven_dP_2022 = Convert.ToDecimal ( dr.Cells["p_2022"].Value);
					Ven_dP_2023 = Convert.ToDecimal ( dr.Cells["p_2023"].Value);
				}

				if ( Convert.ToString( dr.Cells["Venta"].Value) == "TARJETA")
				{
					TVen_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2018"].Value);
					TVen_dP_2019 = Convert.ToDecimal ( dr.Cells["p_2019"].Value);
					TVen_dP_2020 = Convert.ToDecimal ( dr.Cells["p_2020"].Value);
					TVen_dP_2021 = Convert.ToDecimal ( dr.Cells["p_2021"].Value);
					TVen_dP_2022 = Convert.ToDecimal ( dr.Cells["p_2022"].Value);
					TVen_dP_2023 = Convert.ToDecimal ( dr.Cells["p_2023"].Value);
				}

				if ( Convert.ToString( dr.Cells["Venta"].Value) == "CREDITO")
				{
					CVen_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2018"].Value);
					CVen_dP_2019 = Convert.ToDecimal ( dr.Cells["p_2019"].Value);
					CVen_dP_2020 = Convert.ToDecimal ( dr.Cells["p_2020"].Value);
					CVen_dP_2021 = Convert.ToDecimal ( dr.Cells["p_2021"].Value);
					CVen_dP_2022 = Convert.ToDecimal ( dr.Cells["p_2022"].Value);
					CVen_dP_2023 = Convert.ToDecimal ( dr.Cells["p_2023"].Value);
				}

				if ( Convert.ToString( dr.Cells["Venta"].Value) == "DISTRIBUIDOR")
				{
					dVen_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2018"].Value);
					dVen_dP_2019 = Convert.ToDecimal ( dr.Cells["p_2019"].Value);
					dVen_dP_2020 = Convert.ToDecimal ( dr.Cells["p_2020"].Value);
					dVen_dP_2021 = Convert.ToDecimal ( dr.Cells["p_2021"].Value);
					dVen_dP_2022 = Convert.ToDecimal ( dr.Cells["p_2022"].Value);
					dVen_dP_2023 = Convert.ToDecimal ( dr.Cells["p_2023"].Value);
				}

			}	


			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasMargen.Rows.All)
			{
				
				if ( Convert.ToString( dr.Cells["Margen"].Value) == "EFECTIVO")
				{
					Ven_dP_2018 = Ven_dP_2018 - Convert.ToDecimal ( dr.Cells["p_2018"].Value);
					Ven_dP_2019 = Ven_dP_2019 - Convert.ToDecimal ( dr.Cells["p_2019"].Value);
					Ven_dP_2020 = Ven_dP_2020 - Convert.ToDecimal ( dr.Cells["p_2020"].Value);
					Ven_dP_2021 = Ven_dP_2021 - Convert.ToDecimal ( dr.Cells["p_2021"].Value);
					Ven_dP_2022 = Ven_dP_2022 - Convert.ToDecimal ( dr.Cells["p_2022"].Value);
					Ven_dP_2023 = Ven_dP_2023 - Convert.ToDecimal ( dr.Cells["p_2023"].Value);
				}

				if ( Convert.ToString( dr.Cells["Margen"].Value) == "TARJETA")
				{
					TVen_dP_2018 = TVen_dP_2018 - Convert.ToDecimal ( dr.Cells["p_2018"].Value);
					TVen_dP_2019 = TVen_dP_2019 - Convert.ToDecimal ( dr.Cells["p_2019"].Value);
					TVen_dP_2020 = TVen_dP_2020 - Convert.ToDecimal ( dr.Cells["p_2020"].Value);
					TVen_dP_2021 = TVen_dP_2021 - Convert.ToDecimal ( dr.Cells["p_2021"].Value);
					TVen_dP_2022 = TVen_dP_2022 - Convert.ToDecimal ( dr.Cells["p_2022"].Value);
					TVen_dP_2023 = TVen_dP_2023 - Convert.ToDecimal ( dr.Cells["p_2023"].Value);
				}

				if ( Convert.ToString( dr.Cells["Margen"].Value) == "CREDITO")
				{
					CVen_dP_2018 = CVen_dP_2018 - Convert.ToDecimal ( dr.Cells["p_2018"].Value);
					CVen_dP_2019 = CVen_dP_2019 - Convert.ToDecimal ( dr.Cells["p_2019"].Value);
					CVen_dP_2020 = CVen_dP_2020 - Convert.ToDecimal ( dr.Cells["p_2020"].Value);
					CVen_dP_2021 = CVen_dP_2021 - Convert.ToDecimal ( dr.Cells["p_2021"].Value);
					CVen_dP_2022 = CVen_dP_2022 - Convert.ToDecimal ( dr.Cells["p_2022"].Value);
					CVen_dP_2023 = CVen_dP_2023 - Convert.ToDecimal ( dr.Cells["p_2023"].Value);
				}

				if ( Convert.ToString( dr.Cells["Margen"].Value) == "DISTRIBUIDOR")
				{
					dVen_dP_2018 = dVen_dP_2018 - Convert.ToDecimal ( dr.Cells["p_2018"].Value);
					dVen_dP_2019 = dVen_dP_2019 - Convert.ToDecimal ( dr.Cells["p_2019"].Value);
					dVen_dP_2020 = dVen_dP_2020 - Convert.ToDecimal ( dr.Cells["p_2020"].Value);
					dVen_dP_2021 = dVen_dP_2021 - Convert.ToDecimal ( dr.Cells["p_2021"].Value);
					dVen_dP_2022 = dVen_dP_2022 - Convert.ToDecimal ( dr.Cells["p_2022"].Value);
					dVen_dP_2023 = dVen_dP_2023 - Convert.ToDecimal ( dr.Cells["p_2023"].Value);
				}


			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.UGridUb.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Utilidad"].Value) == "EFECTIVO")
				{
					
					dr.Cells["p_2018"].Value =  Math.Round(Ven_dP_2018,2);
					dr.Cells["p_2019"].Value =  Math.Round(Ven_dP_2019,2);
					dr.Cells["p_2020"].Value =  Math.Round(Ven_dP_2020,2);
					dr.Cells["p_2021"].Value =  Math.Round(Ven_dP_2021,2);
					dr.Cells["p_2022"].Value =  Math.Round(Ven_dP_2022,2);
					dr.Cells["p_2023"].Value =  Math.Round(Ven_dP_2023,2);
				
				}
				if ( Convert.ToString( dr.Cells["Utilidad"].Value) == "TARJETA")
				{
					
					dr.Cells["p_2018"].Value =  Math.Round(TVen_dP_2018,2);
					dr.Cells["p_2019"].Value =  Math.Round(TVen_dP_2019,2);
					dr.Cells["p_2020"].Value =  Math.Round(TVen_dP_2020,2);
					dr.Cells["p_2021"].Value =  Math.Round(TVen_dP_2021,2);
					dr.Cells["p_2022"].Value =  Math.Round(TVen_dP_2022,2);
					dr.Cells["p_2023"].Value =  Math.Round(TVen_dP_2023,2);
				
				}
				if ( Convert.ToString( dr.Cells["Utilidad"].Value) == "CREDITO")
				{
					
					dr.Cells["p_2018"].Value =  Math.Round(CVen_dP_2018,2);
					dr.Cells["p_2019"].Value =  Math.Round(CVen_dP_2019,2);
					dr.Cells["p_2020"].Value =  Math.Round(CVen_dP_2020,2);
					dr.Cells["p_2021"].Value =  Math.Round(CVen_dP_2021,2);
					dr.Cells["p_2022"].Value =  Math.Round(CVen_dP_2022,2);
					dr.Cells["p_2023"].Value =  Math.Round(CVen_dP_2023,2);
				
				}
				if ( Convert.ToString( dr.Cells["Utilidad"].Value) == "DISTRIBUIDOR")
				{
					
					dr.Cells["p_2018"].Value =  Math.Round(dVen_dP_2018,2);
					dr.Cells["p_2019"].Value =  Math.Round(dVen_dP_2019,2);
					dr.Cells["p_2020"].Value =  Math.Round(dVen_dP_2020,2);
					dr.Cells["p_2021"].Value =  Math.Round(dVen_dP_2021,2);
					dr.Cells["p_2022"].Value =  Math.Round(dVen_dP_2022,2);
					dr.Cells["p_2023"].Value =  Math.Round(dVen_dP_2023,2);
				
				}
			}

		}

		private void Consulta (int Id)
		{			
			string sSQL = string.Format("Exec CONSULTAGATOSPATRIMONIO {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				if (dr.GetValue(0) != System.DBNull.Value) this.txtGPEnero.Value = dr.GetDecimal(0);
				if (dr.GetValue(1) != System.DBNull.Value) this.txtGPFebrero.Value = dr.GetDecimal(1);
				if (dr.GetValue(2) != System.DBNull.Value) this.txtGPMarzo.Value = dr.GetDecimal(2);
				if (dr.GetValue(3) != System.DBNull.Value) this.txtGPAbril.Value = dr.GetDecimal(3);
				if (dr.GetValue(4) != System.DBNull.Value) this.txtGPMayo.Value = dr.GetDecimal(4);
				if (dr.GetValue(5) != System.DBNull.Value) this.txtGPJunio.Value = dr.GetDecimal(5);
				//				if (dr.GetValue(6) != System.DBNull.Value) this.txtGPJulio.Value = dr.GetDecimal(6);
				//				if (dr.GetValue(7) != System.DBNull.Value) this.txtGPAgosto.Value = dr.GetDecimal(7); 
				//				if (dr.GetValue(8) != System.DBNull.Value) this.txtGPSeptiembre.Value = dr.GetDecimal(8);
				//				if (dr.GetValue(9) != System.DBNull.Value) this.txtGPOctubre.Value = dr.GetDecimal(9);
				//				if (dr.GetValue(10) != System.DBNull.Value) this.txtGPNoviembre.Value = dr.GetDecimal(10);
				//				if (dr.GetValue(11) != System.DBNull.Value) this.txtGPDiciembre.Value = dr.GetDecimal(11);

			}
			dr.Close();


			string sSQLd = string.Format("Exec [CONSULTAGATOSPATRIMONIODETERIORO] {0}", Id);
			SqlDataReader drs = Funcion.rEscalarSQL(cdsSeteoF, sSQLd, true);
			drs.Read();
			if (drs.HasRows)
			{
				if (drs.GetValue(0) != System.DBNull.Value) this.txtdETEnero.Value = drs.GetDecimal(0);
				if (drs.GetValue(1) != System.DBNull.Value) this.txtdETFebrero.Value = drs.GetDecimal(1);
				if (drs.GetValue(2) != System.DBNull.Value) this.txtdETMarzo.Value = drs.GetDecimal(2);
				if (drs.GetValue(3) != System.DBNull.Value) this.txtdETAbril.Value = drs.GetDecimal(3);
				if (drs.GetValue(4) != System.DBNull.Value) this.txtdETMayo.Value = drs.GetDecimal(4);
				if (drs.GetValue(5) != System.DBNull.Value) this.txtdETJunio.Value = drs.GetDecimal(5);
				//				if (dr.GetValue(6) != System.DBNull.Value) this.txtGPJulio.Value = dr.GetDecimal(6);
				//				if (dr.GetValue(7) != System.DBNull.Value) this.txtGPAgosto.Value = dr.GetDecimal(7); 
				//				if (dr.GetValue(8) != System.DBNull.Value) this.txtGPSeptiembre.Value = dr.GetDecimal(8);
				//				if (dr.GetValue(9) != System.DBNull.Value) this.txtGPOctubre.Value = dr.GetDecimal(9);
				//				if (dr.GetValue(10) != System.DBNull.Value) this.txtGPNoviembre.Value = dr.GetDecimal(10);
				//				if (dr.GetValue(11) != System.DBNull.Value) this.txtGPDiciembre.Value = dr.GetDecimal(11);

			}
			drs.Close();
		}

	

		
		private void Resultado ()
		{
			this.txtIngresosEnero.Value = 0.00m;
			this.txtIngresosFebrero.Value = 0.00m;
			this.txtIngresosMarzo.Value = 0.00m;
			this.txtIngresosAbril.Value = 0.00m;
			this.txtIngresosMayo.Value = 0.00m;
			this.txtIngresosJunio.Value = 0.00m;
			this.txtIngresosJulio.Value = 0.00m;
			this.txtIngresosAgosto.Value = 0.00m;
			this.txtIngresosSeptiembre.Value = 0.00m;
			this.txtIngresosOctubre.Value = 0.00m;
			this.txtIngresosNoviembre.Value = 0.00m;
			this.txtIngresosDiciembre.Value = 0.00m;
			this.txtIngresosTotal.Value = 0.00m;
	
			this.txtGastosEnero.Value = 0.00m;
			this.txtGastosFebrero.Value = 0.00m;
			this.txtGastosMarzo.Value = 0.00m;
			this.txtGastosAbril.Value = 0.00m;
			this.txtGastosMayo.Value = 0.00m;
			this.txtGastosJunio.Value = 0.00m;
			this.txtGastosJulio.Value = 0.00m;
			this.txtGastosAgosto.Value = 0.00m;
			this.txtGastosSeptiembre.Value = 0.00m;
			this.txtGastosOctubre.Value = 0.00m;
			this.txtGastosNoviembre.Value = 0.00m;
			this.txtGastosDiciembre.Value = 0.00m;
			this.txtGastosTotal.Value = 0.00m;
	
			this.txtUtilidadEnero.Value = 0.00m;
			this.txtUtilidadFebrero.Value = 0.00m;
			this.txtUtilidadMarzo.Value = 0.00m;
			this.txtUtilidadAbril.Value = 0.00m;
			this.txtUtilidadMayo.Value = 0.00m;
			this.txtUtilidadJunio.Value = 0.00m;
			this.txtUtilidadJulio.Value = 0.00m;
			this.txtUtilidadAgosto.Value = 0.00m;
			this.txtUtilidadSeptiembre.Value = 0.00m;
			this.txtUtilidadOctubre.Value = 0.00m;
			this.txtUtilidadNoviembre.Value = 0.00m;
			this.txtUtilidadDiciembre.Value = 0.00m;
			this.txtUtilidadTotal.Value = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasPrecios.Rows.All)
			{
				this.txtIngresosEnero.Value = Convert.ToDecimal(this.txtIngresosEnero.Value) + Convert.ToDecimal(dr.Cells["P_2018"].Value);
				this.txtIngresosFebrero.Value = Convert.ToDecimal(this.txtIngresosFebrero.Value) + Convert.ToDecimal(dr.Cells["P_2019"].Value);
				this.txtIngresosMarzo.Value = Convert.ToDecimal(this.txtIngresosMarzo.Value) + Convert.ToDecimal(dr.Cells["P_2020"].Value);
				this.txtIngresosAbril.Value = Convert.ToDecimal(this.txtIngresosAbril.Value) + Convert.ToDecimal(dr.Cells["P_2021"].Value);
				this.txtIngresosMayo.Value = Convert.ToDecimal(this.txtIngresosMayo.Value) + Convert.ToDecimal(dr.Cells["P_2022"].Value);
				this.txtIngresosJunio.Value = Convert.ToDecimal(this.txtIngresosJunio.Value) + Convert.ToDecimal(dr.Cells["P_2023"].Value);
				//				this.txtIngresosJulio.Value = Convert.ToDecimal(this.txtIngresosJulio.Value) + Convert.ToDecimal(dr.Cells["JULIO"].Value);
				//				this.txtIngresosAgosto.Value = Convert.ToDecimal(this.txtIngresosAgosto.Value) + Convert.ToDecimal(dr.Cells["AGOSTO"].Value);
				//				this.txtIngresosSeptiembre.Value = Convert.ToDecimal(this.txtIngresosSeptiembre.Value) + Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value);
				//				this.txtIngresosOctubre.Value = Convert.ToDecimal(this.txtIngresosOctubre.Value) + Convert.ToDecimal(dr.Cells["OCTUBRE"].Value);
				//				this.txtIngresosNoviembre.Value = Convert.ToDecimal(this.txtIngresosNoviembre.Value) + Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value);
				//				this.txtIngresosDiciembre.Value = Convert.ToDecimal(this.txtIngresosDiciembre.Value) + Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
				this.txtIngresosTotal.Value = Convert.ToDecimal(this.txtIngresosTotal.Value) + Convert.ToDecimal(dr.Cells["TOTAL"].Value);	
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridOtrosIngresos.Rows.All)
			{
				this.txtIngresosEnero.Value = Convert.ToDecimal(this.txtIngresosEnero.Value) + Convert.ToDecimal(dr.Cells["P_2018"].Value);
				this.txtIngresosFebrero.Value = Convert.ToDecimal(this.txtIngresosFebrero.Value) + Convert.ToDecimal(dr.Cells["P_2019"].Value);
				this.txtIngresosMarzo.Value = Convert.ToDecimal(this.txtIngresosMarzo.Value) + Convert.ToDecimal(dr.Cells["P_2020"].Value);
				this.txtIngresosAbril.Value = Convert.ToDecimal(this.txtIngresosAbril.Value) + Convert.ToDecimal(dr.Cells["P_2021"].Value);
				this.txtIngresosMayo.Value = Convert.ToDecimal(this.txtIngresosMayo.Value) + Convert.ToDecimal(dr.Cells["P_2022"].Value);
				this.txtIngresosJunio.Value = Convert.ToDecimal(this.txtIngresosJunio.Value) + Convert.ToDecimal(dr.Cells["P_2023"].Value);
				//				this.txtIngresosJulio.Value = Convert.ToDecimal(this.txtIngresosJulio.Value) + Convert.ToDecimal(dr.Cells["JULIO"].Value);
				//				this.txtIngresosAgosto.Value = Convert.ToDecimal(this.txtIngresosAgosto.Value) + Convert.ToDecimal(dr.Cells["AGOSTO"].Value);
				//				this.txtIngresosSeptiembre.Value = Convert.ToDecimal(this.txtIngresosSeptiembre.Value) + Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value);
				//				this.txtIngresosOctubre.Value = Convert.ToDecimal(this.txtIngresosOctubre.Value) + Convert.ToDecimal(dr.Cells["OCTUBRE"].Value);
				//				this.txtIngresosNoviembre.Value = Convert.ToDecimal(this.txtIngresosNoviembre.Value) + Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value);
				//				this.txtIngresosDiciembre.Value = Convert.ToDecimal(this.txtIngresosDiciembre.Value) + Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
				this.txtIngresosTotal.Value = Convert.ToDecimal(this.txtIngresosTotal.Value) + Convert.ToDecimal(dr.Cells["TOTAL"].Value);	
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				this.txtGastosEnero.Value = Convert.ToDecimal(this.txtGastosEnero.Value) + Convert.ToDecimal(dr.Cells["P_2018"].Value);
				this.txtGastosFebrero.Value = Convert.ToDecimal(this.txtGastosFebrero.Value) + Convert.ToDecimal(dr.Cells["P_2019"].Value);
				this.txtGastosMarzo.Value = Convert.ToDecimal(this.txtGastosMarzo.Value) + Convert.ToDecimal(dr.Cells["P_2020"].Value);
				this.txtGastosAbril.Value = Convert.ToDecimal(this.txtGastosAbril.Value) + Convert.ToDecimal(dr.Cells["P_2021"].Value);
				this.txtGastosMayo.Value = Convert.ToDecimal(this.txtGastosMayo.Value) + Convert.ToDecimal(dr.Cells["P_2022"].Value);
				this.txtGastosJunio.Value = Convert.ToDecimal(this.txtGastosJunio.Value) + Convert.ToDecimal(dr.Cells["P_2023"].Value);
				//				this.txtGastosJulio.Value = Convert.ToDecimal(this.txtGastosJulio.Value) + Convert.ToDecimal(dr.Cells["JULIO"].Value);
				//				this.txtGastosAgosto.Value = Convert.ToDecimal(this.txtGastosAgosto.Value) + Convert.ToDecimal(dr.Cells["AGOSTO"].Value);
				//				this.txtGastosSeptiembre.Value = Convert.ToDecimal(this.txtGastosSeptiembre.Value) + Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value);
				//				this.txtGastosOctubre.Value = Convert.ToDecimal(this.txtGastosOctubre.Value) + Convert.ToDecimal(dr.Cells["OCTUBRE"].Value);
				//				this.txtGastosNoviembre.Value = Convert.ToDecimal(this.txtGastosNoviembre.Value) + Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value);
				//				this.txtGastosDiciembre.Value = Convert.ToDecimal(this.txtGastosDiciembre.Value) + Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
				this.txtGastosTotal.Value = Convert.ToDecimal(this.txtGastosTotal.Value) + Convert.ToDecimal(dr.Cells["TOTAL"].Value);	
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVentasMargen.Rows.All)
			{
				this.txtGastosEnero.Value = Convert.ToDecimal(this.txtGastosEnero.Value) + Convert.ToDecimal(dr.Cells["P_2018"].Value);
				this.txtGastosFebrero.Value = Convert.ToDecimal(this.txtGastosFebrero.Value) + Convert.ToDecimal(dr.Cells["P_2019"].Value);
				this.txtGastosMarzo.Value = Convert.ToDecimal(this.txtGastosMarzo.Value) + Convert.ToDecimal(dr.Cells["P_2020"].Value);
				this.txtGastosAbril.Value = Convert.ToDecimal(this.txtGastosAbril.Value) + Convert.ToDecimal(dr.Cells["P_2021"].Value);
				this.txtGastosMayo.Value = Convert.ToDecimal(this.txtGastosMayo.Value) + Convert.ToDecimal(dr.Cells["P_2022"].Value);
				this.txtGastosJunio.Value = Convert.ToDecimal(this.txtGastosJunio.Value) + Convert.ToDecimal(dr.Cells["P_2023"].Value);
				//				this.txtGastosJulio.Value = Convert.ToDecimal(this.txtGastosJulio.Value) + Convert.ToDecimal(dr.Cells["JULIO"].Value);
				//				this.txtGastosAgosto.Value = Convert.ToDecimal(this.txtGastosAgosto.Value) + Convert.ToDecimal(dr.Cells["AGOSTO"].Value);
				//				this.txtGastosSeptiembre.Value = Convert.ToDecimal(this.txtGastosSeptiembre.Value) + Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value);
				//				this.txtGastosOctubre.Value = Convert.ToDecimal(this.txtGastosOctubre.Value) + Convert.ToDecimal(dr.Cells["OCTUBRE"].Value);
				//				this.txtGastosNoviembre.Value = Convert.ToDecimal(this.txtGastosNoviembre.Value) + Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value);
				//				this.txtGastosDiciembre.Value = Convert.ToDecimal(this.txtGastosDiciembre.Value) + Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);			
				this.txtGastosTotal.Value = Convert.ToDecimal(this.txtGastosTotal.Value) + Convert.ToDecimal(dr.Cells["TOTAL"].Value);	
			}


			this.txtUtilidadEnero.Value = Convert.ToDecimal(this.txtIngresosEnero.Value) - Convert.ToDecimal(this.txtGastosEnero.Value);
			this.txtUtilidadFebrero.Value = Convert.ToDecimal(this.txtIngresosFebrero.Value) - Convert.ToDecimal(this.txtGastosFebrero.Value);
			this.txtUtilidadMarzo.Value = Convert.ToDecimal(this.txtIngresosMarzo.Value) - Convert.ToDecimal(this.txtGastosMarzo.Value);
			this.txtUtilidadAbril.Value = Convert.ToDecimal(this.txtIngresosAbril.Value) - Convert.ToDecimal(this.txtGastosAbril.Value);
			this.txtUtilidadMayo.Value = Convert.ToDecimal(this.txtIngresosMayo.Value) - Convert.ToDecimal(this.txtGastosMayo.Value);
			this.txtUtilidadJunio.Value = Convert.ToDecimal(this.txtIngresosJunio.Value) - Convert.ToDecimal(this.txtGastosJunio.Value);
			//			this.txtUtilidadJulio.Value = Convert.ToDecimal(this.txtIngresosJulio.Value) - Convert.ToDecimal(this.txtGastosJulio.Value);
			//			this.txtUtilidadAgosto.Value = Convert.ToDecimal(this.txtIngresosAgosto.Value) - Convert.ToDecimal(this.txtGastosAgosto.Value);
			//			this.txtUtilidadSeptiembre.Value = Convert.ToDecimal(this.txtIngresosSeptiembre.Value) - Convert.ToDecimal(this.txtGastosSeptiembre.Value);
			//			this.txtUtilidadOctubre.Value = Convert.ToDecimal(this.txtIngresosOctubre.Value) - Convert.ToDecimal(this.txtGastosOctubre.Value);
			//			this.txtUtilidadNoviembre.Value = Convert.ToDecimal(this.txtIngresosNoviembre.Value) - Convert.ToDecimal(this.txtGastosNoviembre.Value);
			//			this.txtUtilidadDiciembre.Value = Convert.ToDecimal(this.txtIngresosDiciembre.Value) - Convert.ToDecimal(this.txtGastosDiciembre.Value);
			this.txtUtilidadTotal.Value = Convert.ToDecimal(this.txtIngresosTotal.Value) - Convert.ToDecimal(this.txtGastosTotal.Value);	
   
         

			decimal Det_dP_2018 = 0.00m;
			decimal Det_dP_2019 = 0.00m;
			decimal Det_dP_2020 = 0.00m;
			decimal Det_dP_2021 = 0.00m;
			decimal Det_dP_2022 = 0.00m;
			decimal Det_dP_2023 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2018"].Value);
				}
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2019 = Convert.ToDecimal ( dr.Cells["p_2019"].Value);
				}
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2020 = Convert.ToDecimal ( dr.Cells["p_2020"].Value);
				}
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2021 = Convert.ToDecimal ( dr.Cells["p_2021"].Value);
				}
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2022 = Convert.ToDecimal ( dr.Cells["p_2022"].Value);
				}
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2023 = Convert.ToDecimal ( dr.Cells["p_2023"].Value);
				}
			}


			this.txtGP18.Value = Convert.ToDecimal( this.txtGPEnero.Value);
			this.txtGP18T.Value =  Convert.ToDecimal( this.txtGP18.Value) - Convert.ToDecimal( this.txtdETEnero.Value )  ;
   
			this.txtGP19.Value = Convert.ToDecimal( this.txtGPFebrero.Value);
			this.txtGP19T.Value =  Convert.ToDecimal( this.txtGP19.Value) - Convert.ToDecimal( this.txtdETFebrero.Value )  ;

			this.txtGP20.Value = Convert.ToDecimal( this.txtGPMarzo.Value);
			this.txtGP20T.Value =   Convert.ToDecimal( this.txtGP20.Value)- Convert.ToDecimal( this.txtdETMarzo.Value ) ;

			this.txtGP21.Value = Convert.ToDecimal( this.txtGPAbril.Value);
			this.txtGP21T.Value =  Convert.ToDecimal( this.txtGP21.Value)- Convert.ToDecimal( this.txtdETAbril.Value )  ;

			this.txtGP22.Value = Convert.ToDecimal( this.txtGPMayo.Value);
			this.txtGP22T.Value =   Convert.ToDecimal( this.txtGP22.Value)- Convert.ToDecimal( this.txtdETMayo.Value )  ;

			this.txtGP23.Value = Convert.ToDecimal( this.txtGPJunio.Value);
			this.txtGP23T.Value =  Convert.ToDecimal( this.txtGP23.Value) -  Convert.ToDecimal( this.txtdETJunio.Value )  ;


			Totaliza ();


		}

		private void sum_fields(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			
			string sColumna = "";
			decimal dTotal_2018 =  0.00m;
			decimal dTotal_2019 =  0.00m;
			decimal dTotal_2020 =  0.00m;
			decimal dTotal_2021 =  0.00m;
			decimal dTotal_2022 =  0.00m;
			decimal dTotal_2023 =  0.00m;

						
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if((dr.Cells["P_2018"].Value) == System.DBNull.Value ) (dr.Cells["P_2018"].Value) = 0;
				if((dr.Cells["P_2019"].Value) == System.DBNull.Value ) (dr.Cells["P_2019"].Value) = 0;
				if((dr.Cells["P_2020"].Value) == System.DBNull.Value ) (dr.Cells["P_2020"].Value) = 0;
				if((dr.Cells["P_2021"].Value) == System.DBNull.Value ) (dr.Cells["P_2021"].Value) = 0;
				if((dr.Cells["P_2022"].Value) == System.DBNull.Value ) (dr.Cells["P_2022"].Value) = 0;
				if((dr.Cells["P_2023"].Value) == System.DBNull.Value ) (dr.Cells["P_2023"].Value) = 0;
	
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{

				dTotal_2018 = (dTotal_2018 + Convert.ToDecimal(dr.Cells["P_2018"].Value));

				dTotal_2019 = ( dTotal_2019 + Convert.ToDecimal(dr.Cells["P_2019"].Value));

				
				dTotal_2020 = dTotal_2020 + Convert.ToDecimal(dr.Cells["P_2020"].Value);

				
				dTotal_2021 = dTotal_2021 + Convert.ToDecimal(dr.Cells["P_2021"].Value);
				
				
				dTotal_2022 = dTotal_2022 + Convert.ToDecimal(dr.Cells["P_2022"].Value);
				
				
				dTotal_2023 = dTotal_2023 + Convert.ToDecimal(dr.Cells["P_2023"].Value);
	
			}

			this.lblprueba.Text = Convert.ToString( dTotal_2018 );

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{

				if (dTotal_2018 != 0)
				{
					dr.Cells["porcentaje"].Value =  Math.Round((Convert.ToDecimal(dr.Cells["P_2018"].Value) *100)/dTotal_2018,2);
				}
				if (dTotal_2019 !=  0)
				{
					dr.Cells["por_2019"].Value =  Math.Round((Convert.ToDecimal(dr.Cells["P_2019"].Value) *100)/dTotal_2019,2);
				}
				if (dTotal_2020 !=  0)
				{
					dr.Cells["por_2020"].Value =  Math.Round((Convert.ToDecimal(dr.Cells["P_2020"].Value) *100)/dTotal_2020,2);
				}
				if (dTotal_2021 !=  0)
				{
					dr.Cells["por_2021"].Value =  Math.Round((Convert.ToDecimal(dr.Cells["P_2021"].Value) *100)/dTotal_2021,2);
				}
				if (dTotal_2022 !=  0)
				{
					dr.Cells["por_2022"].Value =  Math.Round((Convert.ToDecimal(dr.Cells["P_2022"].Value) *100)/dTotal_2022,2);
				}
				if (dTotal_2023 !=  0)
				{
					dr.Cells["por_2023"].Value =  Math.Round((Convert.ToDecimal(dr.Cells["P_2023"].Value) *100)/dTotal_2023,2);
				}
				/*
				dTotal_2018 = dTotal_2018 + Convert.ToDecimal(dr.Cells["P_2018"].Value);

				dTotal_2019 = dTotal_2019 + Convert.ToDecimal(dr.Cells["P_2019"].Value);

				
				dTotal_2020 = dTotal_2020 + Convert.ToDecimal(dr.Cells["P_2020"].Value);

				
				dTotal_2021 = dTotal_2021 + Convert.ToDecimal(dr.Cells["P_2021"].Value);
				
				
				dTotal_2022 = dTotal_2022 + Convert.ToDecimal(dr.Cells["P_2022"].Value);
				
				
				dTotal_2023 = dTotal_2023 + Convert.ToDecimal(dr.Cells["P_2023"].Value);*/
	
			}


			decimal dTotal =0 ;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				dr.Cells["TOTAL"].Value = 
					Convert.ToDecimal(dr.Cells["P_2018"].Value) + 
					Convert.ToDecimal(dr.Cells["P_2019"].Value) +
					Convert.ToDecimal(dr.Cells["P_2020"].Value) +
					Convert.ToDecimal(dr.Cells["P_2021"].Value) + 
					Convert.ToDecimal(dr.Cells["P_2022"].Value) +
					Convert.ToDecimal(dr.Cells["P_2023"].Value) ;	

				dTotal = dTotal + Convert.ToDecimal( dr.Cells["TOTAL"].Value);
			}

			//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			//			{
			//				if (dTotal>0)
			//				{
			//					dr.Cells["Porcentaje"].Value = (Convert.ToDecimal( dr.Cells["TOTAL"].Value)  * 100 )/dTotal;
			//				}
			//				else
			//				{
			//					dr.Cells["Porcentaje"].Value = 0;
			//				}
			//				 
			//
			//			}

		}
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}


		private void frmPresupuestoReal_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "47002");

			if (!Funcion.Permiso("2027", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Presupuesto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.txtAño.MaxValue = DateTime.Today.Year;
			this.txtAño.Value = DateTime.Today.Year;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
	
		}

		private void uGridVentasPrecios_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "P_2018", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2019", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2020", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2021", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2022", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2023", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "por_2019", 2);
			FuncionesProcedimientos.FormatoGrid(e, "por_2020", 2);
			FuncionesProcedimientos.FormatoGrid(e, "por_2021", 2);
			FuncionesProcedimientos.FormatoGrid(e, "por_2022", 2);
			FuncionesProcedimientos.FormatoGrid(e, "por_2023", 2);

			FuncionesProcedimientos.FormatoGrid(e, "TOTAL", 2);
		

		}

		private void uGridVentasMargen_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "P_2018", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2019", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2020", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2021", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2022", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2023", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "TOTAL", 2);
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
			
			DateTime dtDesde = DateTime.Parse("01/" + sMes + this.txtAño.Value.ToString());
			DateTime dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());

			if ((int)this.uGridGastos.ActiveRow.Cells["idCuentaAgrupa"].Value == 1)
			{
				frmRolReporteDetalleIngresosPorDepartamento rdind = new frmRolReporteDetalleIngresosPorDepartamento(0, dtHasta, false) ;
				rdind.ShowDialog();
			}
			else
			{
				frmPresupuestoResumenGastos prg = new frmPresupuestoResumenGastos(dtDesde, dtHasta, idCuentaAgrupa, sPeriodo, sGasto,0, false, (int)this.txtAño.Value);
				prg.ShowDialog();
			}
		}

		private void uGridGastos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "P_2018", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2019", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2020", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2021", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2022", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2023", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "TOTAL", 2);
		}

		private void btnProyeccion_Click(object sender, System.EventArgs e)
		{
		
		}

		private void uGridOtrosIngresos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "P_2018", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2019", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2020", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2021", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2022", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2023", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "TOTAL", 2);
		}

		private void uGridVentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridVentas);
		}

		private void uGridVentasPrecios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridVentasPrecios);
		}

		private void uGridOtrosIngresos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridOtrosIngresos);
		}

		private void uGridGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridGastos);
	
		}

		private void txtGPEnero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtGPFebrero.Focus();
		}

		private void GuardaPatrimonio (Decimal Valor, int Periodo, int Mes)
		{
			#region Guardar

		
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;

				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					

					#region Graba DetalleMaestro
					
					string  sSQLDet = string.Format("Exec ActualizarPresupuestoReal {0}, {1}, {2}" ,
						Convert.ToDecimal ( Valor), 
						Convert.ToInt32 ( Periodo),
						Convert.ToInt32 ( Mes)
								
						);
					oCmdActualiza.CommandText = sSQLDet;
					oCmdActualiza.ExecuteNonQuery();	
									
					
					#endregion Graba DetalleMaestro
					
					
					oTransaction.Commit();

					Cursor = Cursors.Default;	

				}
				catch(Exception ex)
				{
					try
					{ 
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
				
			#endregion Guardar
		}

		private void GuardaPatrimonioPerdida (Decimal Valor, int Periodo, int Mes)
		{
			#region Guardar

		
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;

				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					

					#region Graba DetalleMaestro
					
					string  sSQLDet = string.Format("Exec ActualizarPresupuestoRealPerdida {0}, {1}, {2}" ,
						Convert.ToDecimal ( Valor), 
						Convert.ToInt32 ( Periodo),
						Convert.ToInt32 ( Mes)
								
						);
					oCmdActualiza.CommandText = sSQLDet;
					oCmdActualiza.ExecuteNonQuery();	
									
					
					#endregion Graba DetalleMaestro
					
					
					oTransaction.Commit();

					Cursor = Cursors.Default;	

				}
				catch(Exception ex)
				{
					try
					{ 
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
				
			#endregion Guardar
		}


		private void txtGPEnero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2018"].Value);
				}
			}

			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPEnero.Value), Convert.ToInt32 (this.txtAño.Value), 1);

			if (this.txtGPEnero.Value != System.DBNull.Value )
			{
				this.txtGP18.Value = Convert.ToDecimal( this.txtGPEnero.Value);
				//this.txtGP18T.Value = Convert.ToDecimal( this.txtGP18.Value) + Convert.ToDecimal( this.txtUtilidadEnero.Value ) + Det_dP_2018 - Convert.ToDecimal( this.txtdETEnero.Value ) ;
				this.txtGP18T.Value =  Convert.ToDecimal( this.txtGP18.Value)  - Convert.ToDecimal( this.txtdETEnero.Value )    ;
				Totaliza ();
			}
		}

		private void Totaliza ()
		{
			this.txtUtilidadTotalP.Value = Convert.ToDecimal(this.txtGP18.Value) +  Convert.ToDecimal(this.txtGP19.Value) +   Convert.ToDecimal(this.txtGP20.Value) + 
				Convert.ToDecimal(this.txtGP21.Value) +  Convert.ToDecimal(this.txtGP22.Value) +  Convert.ToDecimal(this.txtGP23.Value) ;
       
			
			this.txtUtilidadTotalUN.Value = ( Convert.ToDecimal(this.txtGP18T.Value) +  Convert.ToDecimal(this.txtGP19T.Value) +   
				Convert.ToDecimal(this.txtGP20T.Value) + 	Convert.ToDecimal(this.txtGP21T.Value) +   
				Convert.ToDecimal(this.txtGP22T.Value) +   Convert.ToDecimal(this.txtGP23T.Value) );

		}

		private void txtGPEnero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtGPFebrero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2019"].Value);
				}
			}


			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPFebrero.Value), Convert.ToInt32 (this.txtAño.Value), 2);
			if (this.txtGPFebrero.Value != System.DBNull.Value )
			{
				this.txtGP19.Value = Convert.ToDecimal( this.txtGPFebrero.Value);
				//this.txtGP19T.Value = Convert.ToDecimal( this.txtGP19.Value) + Convert.ToDecimal( this.txtUtilidadFebrero.Value )+ Det_dP_2018 -  Convert.ToDecimal( this.txtdETFebrero.Value ) ;
				this.txtGP19T.Value =     Convert.ToDecimal( this.txtGP19.Value) - Convert.ToDecimal( this.txtdETFebrero.Value ) ;
				Totaliza ();
			}
		}

		private void txtGPMarzo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2020"].Value);
				}
			}

			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPMarzo.Value), Convert.ToInt32 (this.txtAño.Value), 3);
			if (this.txtGPMarzo.Value != System.DBNull.Value )
			{
				this.txtGP20.Value = Convert.ToDecimal( this.txtGPMarzo.Value);
				//	this.txtGP20T.Value = Convert.ToDecimal( this.txtGP20.Value) + Convert.ToDecimal( this.txtUtilidadMarzo.Value ) + Det_dP_2018 - Convert.ToDecimal( this.txtdETMarzo.Value ) ;;
				this.txtGP20T.Value =  Convert.ToDecimal( this.txtGP20.Value) - Convert.ToDecimal( this.txtdETMarzo.Value )  ;
				Totaliza ();
			}

		}

		private void txtGPAbril_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2021"].Value);
				}
			}


			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPAbril.Value), Convert.ToInt32 (this.txtAño.Value), 4);
			if (this.txtGPAbril.Value != System.DBNull.Value )
			{
				this.txtGP21.Value = Convert.ToDecimal( this.txtGPAbril.Value);
				//this.txtGP21T.Value = Convert.ToDecimal( this.txtGP21.Value) + Convert.ToDecimal( this.txtUtilidadAbril.Value )+ Det_dP_2018 - Convert.ToDecimal( this.txtdETAbril.Value ) ;;
				this.txtGP21T.Value =   Convert.ToDecimal( this.txtGP21.Value) - Convert.ToDecimal( this.txtdETAbril.Value )  ;
				Totaliza ();
			}

		}

		private void txtGPMayo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2022"].Value);
				}
			}

			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPMayo.Value), Convert.ToInt32 (this.txtAño.Value), 5);
			if (this.txtGPMayo.Value != System.DBNull.Value )
			{
				this.txtGP22.Value = Convert.ToDecimal( this.txtGPMayo.Value);
				//this.txtGP22T.Value = Convert.ToDecimal( this.txtGP21.Value) + Convert.ToDecimal( this.txtUtilidadMayo.Value )+ Det_dP_2018 - Convert.ToDecimal( this.txtdETMayo.Value ) ;;
				this.txtGP22T.Value =   Convert.ToDecimal( this.txtGP21.Value) - Convert.ToDecimal( this.txtdETMayo.Value ) ;
				Totaliza ();
			}


		}

		private void txtGPJunio_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2023"].Value);
				}
			}

			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPJunio.Value), Convert.ToInt32 (this.txtAño.Value), 6);
			if (this.txtGPJunio.Value != System.DBNull.Value )
			{
				this.txtGP23.Value = Convert.ToDecimal( this.txtGPJunio.Value);
				//this.txtGP23T.Value = Convert.ToDecimal( this.txtGP23.Value) + Convert.ToDecimal( this.txtUtilidadJunio.Value ) + Det_dP_2018 - Convert.ToDecimal( this.txtdETJunio.Value ) ;;
				this.txtGP23T.Value =  Convert.ToDecimal( this.txtGP23.Value) - Convert.ToDecimal( this.txtdETJunio.Value )  ;
				Totaliza ();
			}

		}

		private void txtGPJulio_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPJulio.Value), Convert.ToInt32 (this.txtAño.Value), 7);
		}

		private void txtGPAgosto_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPAgosto.Value), Convert.ToInt32 (this.txtAño.Value), 8);
		}

		private void txtGPSeptiembre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPSeptiembre.Value), Convert.ToInt32 (this.txtAño.Value), 9);
		}

		private void txtGPOctubre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPOctubre.Value), Convert.ToInt32 (this.txtAño.Value), 10);
		}

		private void txtGPNoviembre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPNoviembre.Value), Convert.ToInt32 (this.txtAño.Value), 11);
		}

		private void txtGPDiciembre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			this.GuardaPatrimonio(Convert.ToDecimal( this.txtGPDiciembre.Value), Convert.ToInt32 (this.txtAño.Value), 12);
		}

		private void txtGPEnero_Click(object sender, System.EventArgs e)
		{
			this.txtGPEnero.SelectAll();
		}

		private void txtGPFebrero_Click(object sender, System.EventArgs e)
		{
			this.txtGPFebrero.SelectAll();
		}

		private void txtGPMarzo_Click(object sender, System.EventArgs e)
		{
			this.txtGPMarzo.SelectAll();
		}

		private void txtGPAbril_Click(object sender, System.EventArgs e)
		{
			this.txtGPAbril.SelectAll();
		}

		private void txtGPMayo_Click(object sender, System.EventArgs e)
		{
			this.txtGPMayo.SelectAll();
		}

		private void txtGPJunio_Click(object sender, System.EventArgs e)
		{
			this.txtGPJunio.SelectAll();
		}

		private void txtGPJulio_Click(object sender, System.EventArgs e)
		{
			this.txtGPJulio.SelectAll();
			
		}

		private void txtGPAgosto_Click(object sender, System.EventArgs e)
		{
			this.txtGPAgosto.SelectAll();
		}

		private void txtGPSeptiembre_Click(object sender, System.EventArgs e)
		{
			this.txtGPSeptiembre.SelectAll();
		}

		private void txtGPOctubre_Click(object sender, System.EventArgs e)
		{
			this.txtGPOctubre.SelectAll();
		}

		private void txtGPNoviembre_Click(object sender, System.EventArgs e)
		{
			this.txtGPNoviembre.SelectAll();
		}

		private void txtGPDiciembre_Click(object sender, System.EventArgs e)
		{
			this.txtGPDiciembre.SelectAll();
		}

		private void txtGPFebrero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)	this.txtGPMarzo.Focus();
		}

		private void txtGPMarzo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)	this.txtGPAbril.Focus();
		}

		private void txtGPAbril_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)	this.txtGPMayo.Focus();
		}

		private void txtGPMayo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)	this.txtGPJunio.Focus();
		}

		private void txtGPJunio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.txtGPJulio.Focus();
		}

		private void txtGPJulio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)	this.txtGPAgosto.Focus();
		}

		private void txtGPAgosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)	this.txtGPSeptiembre.Focus();
		}

		private void txtGPSeptiembre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.txtGPOctubre.Focus();
		}

		private void txtGPOctubre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.txtGPNoviembre.Focus();
		}

		private void txtGPNoviembre_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void txtGPNoviembre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.txtGPDiciembre.Focus();
		}

		private void frmPresupuestoReal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void label10_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label17_Click(object sender, System.EventArgs e)
		{
		
		}

		private void UGridUb_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "P_2018", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2019", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2020", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2021", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2022", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2023", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "TOTAL", 2);
		}

		private void txtdETEnero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2018"].Value);
				}
			}

			this.GuardaPatrimonioPerdida(Convert.ToDecimal( this.txtdETEnero.Value), Convert.ToInt32 (this.txtAño.Value), 1);

			if (this.txtdETEnero.Value != System.DBNull.Value )
			{
				this.txtGP18.Value = Convert.ToDecimal( this.txtGPEnero.Value);
				//this.txtGP18T.Value = Convert.ToDecimal( this.txtGP18.Value) + Convert.ToDecimal( this.txtUtilidadEnero.Value ) + Det_dP_2018 - Convert.ToDecimal( this.txtdETEnero.Value ) ;
				//this.txtGP18T.Value = Convert.ToDecimal( this.txtGP18.Value) + Convert.ToDecimal( this.txtUtilidadEnero.Value ) + Det_dP_2018 - Convert.ToDecimal( this.txtdETEnero.Value ) ;
				this.txtGP18T.Value =   Convert.ToDecimal( this.txtGP18.Value)- Convert.ToDecimal( this.txtdETEnero.Value ) ;
				Totaliza ();
			}
		}

		private void txtdETFebrero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2019"].Value);
				}
			}

			this.GuardaPatrimonioPerdida(Convert.ToDecimal( this.txtdETFebrero.Value), Convert.ToInt32 (this.txtAño.Value), 2);

			if (this.txtdETFebrero.Value != System.DBNull.Value )
			{
				this.txtGP19.Value = Convert.ToDecimal( this.txtGPFebrero.Value);
				//this.txtGP19T.Value = Convert.ToDecimal( this.txtGP19.Value) + Convert.ToDecimal( this.txtUtilidadFebrero.Value ) +Det_dP_2018 - Convert.ToDecimal( this.txtdETFebrero.Value ) ;
				this.txtGP19T.Value = Convert.ToDecimal( this.txtGP19.Value) -  Convert.ToDecimal( this.txtdETFebrero.Value )   ;
				Totaliza ();
			}
		}

		private void txtdETMarzo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2020"].Value);
				}
			}

			this.GuardaPatrimonioPerdida(Convert.ToDecimal( this.txtdETMarzo.Value), Convert.ToInt32 (this.txtAño.Value), 3);

			if (this.txtdETMarzo.Value != System.DBNull.Value )
			{
				this.txtGP20.Value = Convert.ToDecimal( this.txtGPMarzo.Value);
				//this.txtGP20T.Value = Convert.ToDecimal( this.txtGP20.Value) + Convert.ToDecimal( this.txtUtilidadMarzo.Value )+ Det_dP_2018   - Convert.ToDecimal( this.txtdETMarzo.Value ) ;
				this.txtGP20T.Value = Convert.ToDecimal( this.txtGP20.Value) - Convert.ToDecimal( this.txtdETMarzo.Value )  ;
				Totaliza ();
			}
		}

		private void txtdETAbril_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2021"].Value);
				}
			}

			this.GuardaPatrimonioPerdida(Convert.ToDecimal( this.txtdETAbril.Value), Convert.ToInt32 (this.txtAño.Value), 4);

			if (this.txtdETAbril.Value != System.DBNull.Value )
			{
				this.txtGP21.Value = Convert.ToDecimal( this.txtGPAbril.Value);
				//this.txtGP21T.Value = Convert.ToDecimal( this.txtGP21.Value) + Convert.ToDecimal( this.txtUtilidadAbril.Value ) + Det_dP_2018  - Convert.ToDecimal( this.txtdETAbril.Value ) ;
				this.txtGP21T.Value =     Convert.ToDecimal( this.txtGP21.Value)- Convert.ToDecimal( this.txtdETAbril.Value )  ;
				Totaliza ();
			}
		}

		private void txtdETMayo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2022"].Value);
				}
			}
			this.GuardaPatrimonioPerdida(Convert.ToDecimal( this.txtdETMayo.Value), Convert.ToInt32 (this.txtAño.Value), 5);

			if (this.txtdETMayo.Value != System.DBNull.Value )
			{
				this.txtGP22.Value = Convert.ToDecimal( this.txtGPMayo.Value);
				//this.txtGP22T.Value = Convert.ToDecimal( this.txtGP22.Value) + Convert.ToDecimal( this.txtUtilidadMayo.Value ) + Det_dP_2018  - Convert.ToDecimal( this.txtdETMayo.Value ) ;
				this.txtGP22T.Value =   Convert.ToDecimal( this.txtGP22.Value)- Convert.ToDecimal( this.txtdETMayo.Value ) ;
				Totaliza ();
			}
		}

		private void txtdETJunio_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			decimal Det_dP_2018 = 0.00m;
			string sColumna = "";
			int iContColumnas = (int)this.uGridGastos.DisplayLayout.Bands[0].Columns.Count;	

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridGastos.Rows.All)
			{
				if ( Convert.ToString( dr.Cells["Grupo"].Value) == "PERDIDAS Y DETERIORO")
				{
					Det_dP_2018 = Convert.ToDecimal ( dr.Cells["p_2023"].Value);
				}
			}
			this.GuardaPatrimonioPerdida(Convert.ToDecimal( this.txtdETJunio.Value), Convert.ToInt32 (this.txtAño.Value), 6);

			if (this.txtdETJunio.Value != System.DBNull.Value )
			{
				this.txtGP23.Value = Convert.ToDecimal( this.txtGPJunio.Value);
				//this.txtGP23T.Value = Convert.ToDecimal( this.txtGP23.Value) + Convert.ToDecimal( this.txtUtilidadJunio.Value ) + Det_dP_2018 - Convert.ToDecimal( this.txtdETJunio.Value ) ;
				this.txtGP23T.Value =  Convert.ToDecimal( this.txtGP23.Value)- Convert.ToDecimal( this.txtdETJunio.Value )   ;
				Totaliza ();
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "P_2018", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2019", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2020", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2021", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2022", 2);
			FuncionesProcedimientos.FormatoGrid(e, "P_2023", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "TOTAL", 2);
		}
	}
}
