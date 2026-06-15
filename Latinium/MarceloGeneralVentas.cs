using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de MarceloGeneralVentas.
	/// </summary>
	public class MarceloGeneralVentas : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label29;
		private System.Windows.Forms.Label label30;
		private System.Windows.Forms.Label label31;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.Label label34;
		private System.Windows.Forms.Label label35;
		private System.Windows.Forms.Label label36;
		private System.Windows.Forms.Label label37;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsMarcelo;
		private C1.Data.C1DataView cdvMarcelo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEfectivo;
		private System.Windows.Forms.Label label38;
		private System.Windows.Forms.Label label39;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaFacil;
		private System.Windows.Forms.Label label40;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtOtrasTarjetas;
		private System.Windows.Forms.Label label41;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastos;
		private System.Windows.Forms.Label label42;
		private System.Windows.Forms.Label label43;
		private System.Windows.Forms.Label label44;
		private System.Windows.Forms.Label label45;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDeposito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCredPersonal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalTarjetas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCheque;
		private System.Windows.Forms.Label label46;
		private System.Windows.Forms.Label label47;
		private System.Windows.Forms.Label label48;
		private System.Windows.Forms.Label label49;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorActual;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUnidadActual;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorAntes;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUnidadesAntes;
		private System.Windows.Forms.Label label50;
		private System.Windows.Forms.Label label51;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label label54;
		private System.Windows.Forms.Label label55;
		private System.Windows.Forms.Label label56;
		private System.Windows.Forms.Label label57;
		private System.Windows.Forms.Label label58;
		private System.Windows.Forms.Label label59;
		private System.Windows.Forms.Label label60;
		private System.Windows.Forms.Label label61;
		private System.Windows.Forms.Label label62;
		private System.Windows.Forms.Label label63;
		private System.Windows.Forms.Label label64;
		private System.Windows.Forms.Label label65;
		private System.Windows.Forms.Label label66;
		private System.Windows.Forms.Label label67;
		private System.Windows.Forms.Label label68;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHombre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMujer;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNiño;
		private System.Windows.Forms.Label lblfecha;
		private C1.Data.C1DataSet cdsMarceloBodega;
		private C1.Data.C1DataView cdvMarceloBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPromo29;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPromo15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPromo10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPromo5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBalon;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPortazapatos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBilletera;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCanguro;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMaleta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMochila;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGorra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCinturon;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCartera;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNiña;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPromoOtro;
		private System.Windows.Forms.Label label69;
		private System.Windows.Forms.Label label52;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPromo20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPromo25;
		private System.Windows.Forms.Label label70;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private System.Windows.Forms.Button btGrabar;
		private System.Windows.Forms.Button btDeshacer;
		private System.Windows.Forms.Button btReportes;
		private System.Windows.Forms.Label label71;
		private System.Windows.Forms.Label label73;
		private System.Windows.Forms.Label label74;
		private System.Windows.Forms.Label label75;
		private System.Windows.Forms.Label label76;
		private System.Windows.Forms.Label label77;
		private System.Windows.Forms.Label label78;
		private System.Windows.Forms.Label label79;
		private System.Windows.Forms.Label label80;
		private System.Windows.Forms.Label label81;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor7;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor8;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor14;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor16;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor17;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor19;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor22;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor23;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor24;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor25;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor26;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor27;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor28;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor29;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor30;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor31;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor33;
		private System.Windows.Forms.Label label83;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor34;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor38;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor39;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor40;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor41;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor42;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor45;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor49;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor50;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor51;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor52;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor53;
		private System.Windows.Forms.Label label86;
		private System.Windows.Forms.Label label87;
		private System.Windows.Forms.Label label88;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIniciaF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTerminaF;
		private System.Windows.Forms.Label label91;
		private System.Windows.Forms.Label label92;
		private System.Windows.Forms.Label label93;
		private System.Windows.Forms.Label label94;
		private System.Windows.Forms.Label label95;
		private System.Windows.Forms.Label label96;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPromo35;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPromo79;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetencion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCredEmpresa;
		private System.Windows.Forms.Label label97;
		private System.Windows.Forms.Label label98;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTerminaNC;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIniciaNC;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor48;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor37;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor36;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor47;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor46;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor35;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor21;
		private C1.Data.C1DataSet cdsFormaPago;
		private C1.Data.C1DataView cdvFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo1;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo2;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo3;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo4;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo5;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo6;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo7;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo8;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo9;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo10;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor32;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor43;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor44;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor54;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor55;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor56;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor57;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor58;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo14;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor59;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor60;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor61;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor62;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo16;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor63;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor64;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo17;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor65;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor66;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor67;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor68;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo19;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor69;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor70;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo20;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor cefact1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor cefact2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor cefact3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor cefact4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor cet1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor cet2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor cet3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor cet4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor cet5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor cefact5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre5;
		private System.Windows.Forms.Button btTraer;
		private System.Windows.Forms.Label label72;
		private System.Windows.Forms.Label label82;
		private System.Windows.Forms.Label label84;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor71;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor72;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor73;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor74;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor75;
		private System.Windows.Forms.Label label85;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MarceloGeneralVentas()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
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
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand19 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand20 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand21 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
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
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.label24 = new System.Windows.Forms.Label();
			this.label25 = new System.Windows.Forms.Label();
			this.label26 = new System.Windows.Forms.Label();
			this.label27 = new System.Windows.Forms.Label();
			this.label28 = new System.Windows.Forms.Label();
			this.label29 = new System.Windows.Forms.Label();
			this.label30 = new System.Windows.Forms.Label();
			this.label31 = new System.Windows.Forms.Label();
			this.label32 = new System.Windows.Forms.Label();
			this.label33 = new System.Windows.Forms.Label();
			this.label34 = new System.Windows.Forms.Label();
			this.label35 = new System.Windows.Forms.Label();
			this.label36 = new System.Windows.Forms.Label();
			this.label37 = new System.Windows.Forms.Label();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsMarcelo = new C1.Data.C1DataSet();
			this.cdvMarcelo = new C1.Data.C1DataView();
			this.txtEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label38 = new System.Windows.Forms.Label();
			this.label39 = new System.Windows.Forms.Label();
			this.txtCuotaFacil = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label40 = new System.Windows.Forms.Label();
			this.txtOtrasTarjetas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label41 = new System.Windows.Forms.Label();
			this.txtGastos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label42 = new System.Windows.Forms.Label();
			this.txtDeposito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label43 = new System.Windows.Forms.Label();
			this.txtCredPersonal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label44 = new System.Windows.Forms.Label();
			this.txtTotalTarjetas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label45 = new System.Windows.Forms.Label();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label46 = new System.Windows.Forms.Label();
			this.label47 = new System.Windows.Forms.Label();
			this.label48 = new System.Windows.Forms.Label();
			this.label49 = new System.Windows.Forms.Label();
			this.txtValorActual = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUnidadActual = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtValorAntes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtUnidadesAntes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label50 = new System.Windows.Forms.Label();
			this.label51 = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.label54 = new System.Windows.Forms.Label();
			this.label55 = new System.Windows.Forms.Label();
			this.label56 = new System.Windows.Forms.Label();
			this.label57 = new System.Windows.Forms.Label();
			this.label58 = new System.Windows.Forms.Label();
			this.label59 = new System.Windows.Forms.Label();
			this.label60 = new System.Windows.Forms.Label();
			this.label61 = new System.Windows.Forms.Label();
			this.label62 = new System.Windows.Forms.Label();
			this.label63 = new System.Windows.Forms.Label();
			this.label64 = new System.Windows.Forms.Label();
			this.label65 = new System.Windows.Forms.Label();
			this.label66 = new System.Windows.Forms.Label();
			this.label67 = new System.Windows.Forms.Label();
			this.label68 = new System.Windows.Forms.Label();
			this.txtHombre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMujer = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNiño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblfecha = new System.Windows.Forms.Label();
			this.cdsMarceloBodega = new C1.Data.C1DataSet();
			this.cdvMarceloBodega = new C1.Data.C1DataView();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtPromo29 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPromo15 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPromo10 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPromo5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBalon = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPortazapatos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBilletera = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCanguro = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMaleta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMochila = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGorra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCinturon = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCartera = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNiña = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPromoOtro = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label69 = new System.Windows.Forms.Label();
			this.label52 = new System.Windows.Forms.Label();
			this.txtPromo20 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPromo25 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label70 = new System.Windows.Forms.Label();
			this.btGrabar = new System.Windows.Forms.Button();
			this.btDeshacer = new System.Windows.Forms.Button();
			this.btReportes = new System.Windows.Forms.Button();
			this.label71 = new System.Windows.Forms.Label();
			this.label73 = new System.Windows.Forms.Label();
			this.label74 = new System.Windows.Forms.Label();
			this.label75 = new System.Windows.Forms.Label();
			this.label76 = new System.Windows.Forms.Label();
			this.label77 = new System.Windows.Forms.Label();
			this.label78 = new System.Windows.Forms.Label();
			this.label79 = new System.Windows.Forms.Label();
			this.label80 = new System.Windows.Forms.Label();
			this.label81 = new System.Windows.Forms.Label();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor6 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor7 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor8 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor9 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor10 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor11 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor12 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor13 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor14 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor15 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor16 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor17 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor18 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor19 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor20 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor22 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor23 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor24 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor25 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor26 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor27 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor28 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor29 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor30 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor31 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor33 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label83 = new System.Windows.Forms.Label();
			this.ultraNumericEditor34 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor38 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor39 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor40 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor41 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor42 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor45 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor49 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor50 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor51 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor52 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor53 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label86 = new System.Windows.Forms.Label();
			this.label87 = new System.Windows.Forms.Label();
			this.label88 = new System.Windows.Forms.Label();
			this.txtIniciaF = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTerminaF = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label91 = new System.Windows.Forms.Label();
			this.label92 = new System.Windows.Forms.Label();
			this.label93 = new System.Windows.Forms.Label();
			this.label94 = new System.Windows.Forms.Label();
			this.label95 = new System.Windows.Forms.Label();
			this.label96 = new System.Windows.Forms.Label();
			this.txtPromo35 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPromo79 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtRetencion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCredEmpresa = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label97 = new System.Windows.Forms.Label();
			this.label98 = new System.Windows.Forms.Label();
			this.txtTerminaNC = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIniciaNC = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor48 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor37 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor36 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor47 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor46 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor35 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor21 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsFormaPago = new C1.Data.C1DataSet();
			this.cdvFormaPago = new C1.Data.C1DataView();
			this.ultraCombo1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo3 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo4 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo5 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo6 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo7 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo8 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo9 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo10 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo11 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraNumericEditor32 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor43 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor44 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor54 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCombo12 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraNumericEditor55 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor56 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCombo13 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraNumericEditor57 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor58 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCombo14 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraNumericEditor59 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor60 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCombo15 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraNumericEditor61 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor62 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCombo16 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraNumericEditor63 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor64 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCombo17 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraNumericEditor65 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor66 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCombo18 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraNumericEditor67 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor68 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCombo19 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraNumericEditor69 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor70 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraCombo20 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNombre1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cefact1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cefact2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cefact3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cefact4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cet1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cet2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cet3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cet4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cet5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cefact5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNombre5 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btTraer = new System.Windows.Forms.Button();
			this.label72 = new System.Windows.Forms.Label();
			this.label82 = new System.Windows.Forms.Label();
			this.label84 = new System.Windows.Forms.Label();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor5 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraNumericEditor71 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor72 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor73 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor74 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor75 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label85 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarcelo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarcelo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaFacil)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtrasTarjetas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalTarjetas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorActual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUnidadActual)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorAntes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUnidadesAntes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMujer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNiño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarceloBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo29)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo15)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBalon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPortazapatos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBilletera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCanguro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMaleta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMochila)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGorra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCinturon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCartera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNiña)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromoOtro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo25)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor15)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor16)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor17)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor18)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor19)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor22)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor23)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor24)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor25)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor26)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor27)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor28)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor29)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor30)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor31)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor33)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor34)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor38)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor39)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor40)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor41)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor42)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor45)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor49)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor50)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor51)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor52)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor53)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIniciaF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTerminaF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo35)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo79)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTerminaNC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIniciaNC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor48)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor37)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor36)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor47)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor46)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor35)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor21)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor32)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor43)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor44)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor54)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor55)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor56)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor57)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor58)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor59)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor60)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo15)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor61)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor62)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo16)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor63)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor64)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo17)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor65)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor66)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo18)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor67)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor68)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo19)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor69)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor70)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cefact1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cefact2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cefact3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cefact4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cet1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cet2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cet3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cet4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cet5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cefact5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor71)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor72)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor73)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor74)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor75)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(0, 0);
			this.label1.Name = "label1";
			this.label1.TabIndex = 0;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(0, 0);
			this.label2.Name = "label2";
			this.label2.TabIndex = 0;
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(0, 0);
			this.label3.Name = "label3";
			this.label3.TabIndex = 0;
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(0, 0);
			this.label4.Name = "label4";
			this.label4.TabIndex = 0;
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(0, 0);
			this.label5.Name = "label5";
			this.label5.TabIndex = 0;
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(0, 0);
			this.label6.Name = "label6";
			this.label6.TabIndex = 0;
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(0, 0);
			this.label7.Name = "label7";
			this.label7.TabIndex = 0;
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(0, 0);
			this.label8.Name = "label8";
			this.label8.TabIndex = 0;
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(0, 0);
			this.label9.Name = "label9";
			this.label9.TabIndex = 0;
			// 
			// label10
			// 
			this.label10.Location = new System.Drawing.Point(0, 0);
			this.label10.Name = "label10";
			this.label10.TabIndex = 0;
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(0, 0);
			this.label11.Name = "label11";
			this.label11.TabIndex = 0;
			// 
			// label12
			// 
			this.label12.Location = new System.Drawing.Point(0, 0);
			this.label12.Name = "label12";
			this.label12.TabIndex = 0;
			// 
			// label13
			// 
			this.label13.Location = new System.Drawing.Point(0, 0);
			this.label13.Name = "label13";
			this.label13.TabIndex = 0;
			// 
			// label14
			// 
			this.label14.Location = new System.Drawing.Point(0, 0);
			this.label14.Name = "label14";
			this.label14.TabIndex = 0;
			// 
			// label15
			// 
			this.label15.Location = new System.Drawing.Point(0, 0);
			this.label15.Name = "label15";
			this.label15.TabIndex = 0;
			// 
			// label16
			// 
			this.label16.Location = new System.Drawing.Point(0, 0);
			this.label16.Name = "label16";
			this.label16.TabIndex = 0;
			// 
			// label17
			// 
			this.label17.Location = new System.Drawing.Point(0, 0);
			this.label17.Name = "label17";
			this.label17.TabIndex = 0;
			// 
			// label18
			// 
			this.label18.Location = new System.Drawing.Point(0, 0);
			this.label18.Name = "label18";
			this.label18.TabIndex = 0;
			// 
			// label19
			// 
			this.label19.Location = new System.Drawing.Point(0, 0);
			this.label19.Name = "label19";
			this.label19.TabIndex = 0;
			// 
			// label20
			// 
			this.label20.Location = new System.Drawing.Point(0, 0);
			this.label20.Name = "label20";
			this.label20.TabIndex = 0;
			// 
			// label21
			// 
			this.label21.Location = new System.Drawing.Point(0, 0);
			this.label21.Name = "label21";
			this.label21.TabIndex = 0;
			// 
			// label22
			// 
			this.label22.Location = new System.Drawing.Point(0, 0);
			this.label22.Name = "label22";
			this.label22.TabIndex = 0;
			// 
			// label23
			// 
			this.label23.Location = new System.Drawing.Point(0, 0);
			this.label23.Name = "label23";
			this.label23.TabIndex = 0;
			// 
			// label24
			// 
			this.label24.Location = new System.Drawing.Point(0, 0);
			this.label24.Name = "label24";
			this.label24.TabIndex = 0;
			// 
			// label25
			// 
			this.label25.Location = new System.Drawing.Point(0, 0);
			this.label25.Name = "label25";
			this.label25.TabIndex = 0;
			// 
			// label26
			// 
			this.label26.Location = new System.Drawing.Point(0, 0);
			this.label26.Name = "label26";
			this.label26.TabIndex = 0;
			// 
			// label27
			// 
			this.label27.Location = new System.Drawing.Point(0, 0);
			this.label27.Name = "label27";
			this.label27.TabIndex = 0;
			// 
			// label28
			// 
			this.label28.Location = new System.Drawing.Point(0, 0);
			this.label28.Name = "label28";
			this.label28.TabIndex = 0;
			// 
			// label29
			// 
			this.label29.Location = new System.Drawing.Point(0, 0);
			this.label29.Name = "label29";
			this.label29.TabIndex = 0;
			// 
			// label30
			// 
			this.label30.Location = new System.Drawing.Point(0, 0);
			this.label30.Name = "label30";
			this.label30.TabIndex = 0;
			// 
			// label31
			// 
			this.label31.Location = new System.Drawing.Point(0, 0);
			this.label31.Name = "label31";
			this.label31.TabIndex = 0;
			// 
			// label32
			// 
			this.label32.Location = new System.Drawing.Point(0, 0);
			this.label32.Name = "label32";
			this.label32.TabIndex = 0;
			// 
			// label33
			// 
			this.label33.Location = new System.Drawing.Point(0, 0);
			this.label33.Name = "label33";
			this.label33.TabIndex = 0;
			// 
			// label34
			// 
			this.label34.Location = new System.Drawing.Point(0, 0);
			this.label34.Name = "label34";
			this.label34.TabIndex = 0;
			// 
			// label35
			// 
			this.label35.Location = new System.Drawing.Point(0, 0);
			this.label35.Name = "label35";
			this.label35.TabIndex = 0;
			// 
			// label36
			// 
			this.label36.Location = new System.Drawing.Point(0, 0);
			this.label36.Name = "label36";
			this.label36.TabIndex = 0;
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label37.Location = new System.Drawing.Point(576, 8);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(325, 27);
			this.label37.TabIndex = 1;
			this.label37.Text = "MARCELO GENERAL VENTAS";
			this.label37.Click += new System.EventHandler(this.label37_Click);
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "MarceloGeneralVentas";
			this.barraDato1.DataNombreId = "idMarceloGeneralVentas";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsMarcelo;
			this.barraDato1.DataTabla = "MarceloGeneralVentas";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Enabled = false;
			this.barraDato1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.barraDato1.Location = new System.Drawing.Point(672, 32);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 0;
			this.barraDato1.Visible = false;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cdsMarcelo
			// 
			this.cdsMarcelo.BindingContextCtrl = this;
			this.cdsMarcelo.DataLibrary = "LibFacturacion";
			this.cdsMarcelo.DataLibraryUrl = "";
			this.cdsMarcelo.DataSetDef = "dsMarceloGeneralVentas";
			this.cdsMarcelo.FillOnRequest = false;
			this.cdsMarcelo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsMarcelo.Name = "cdsMarcelo";
			this.cdsMarcelo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMarcelo.SchemaDef = null;
			this.cdsMarcelo.BeforeFill += new C1.Data.FillEventHandler(this.cdsMarcelo_BeforeFill);
			// 
			// cdvMarcelo
			// 
			this.cdvMarcelo.BindingContextCtrl = this;
			this.cdvMarcelo.DataSet = this.cdsMarcelo;
			this.cdvMarcelo.Sort = "clave";
			this.cdvMarcelo.TableName = "";
			this.cdvMarcelo.TableViewName = "MarceloGeneralVentas";
			// 
			// txtEfectivo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEfectivo.Appearance = appearance1;
			this.txtEfectivo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Efectivo"));
			this.txtEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEfectivo.FormatString = "#,##0.00";
			this.txtEfectivo.Location = new System.Drawing.Point(80, 40);
			this.txtEfectivo.Name = "txtEfectivo";
			this.txtEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtEfectivo.PromptChar = ' ';
			this.txtEfectivo.Size = new System.Drawing.Size(104, 21);
			this.txtEfectivo.TabIndex = 2;
			this.txtEfectivo.Validated += new System.EventHandler(this.txtEfectivo_Validated);
			this.txtEfectivo.ValueChanged += new System.EventHandler(this.txtEfectivo_ValueChanged);
			this.txtEfectivo.Enter += new System.EventHandler(this.txtEfectivo_Enter);
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.Location = new System.Drawing.Point(16, 48);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(45, 16);
			this.label38.TabIndex = 9;
			this.label38.Text = "Efectivo";
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.Location = new System.Drawing.Point(16, 72);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(61, 16);
			this.label39.TabIndex = 11;
			this.label39.Text = "Cuota Facil";
			// 
			// txtCuotaFacil
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaFacil.Appearance = appearance2;
			this.txtCuotaFacil.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "CoutaFacil"));
			this.txtCuotaFacil.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaFacil.FormatString = "#,##0.00";
			this.txtCuotaFacil.Location = new System.Drawing.Point(80, 64);
			this.txtCuotaFacil.Name = "txtCuotaFacil";
			this.txtCuotaFacil.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaFacil.PromptChar = ' ';
			this.txtCuotaFacil.Size = new System.Drawing.Size(104, 21);
			this.txtCuotaFacil.TabIndex = 4;
			this.txtCuotaFacil.Validated += new System.EventHandler(this.txtCuotaFacil_Validated);
			this.txtCuotaFacil.ValueChanged += new System.EventHandler(this.txtCuotaFacil_ValueChanged);
			this.txtCuotaFacil.Enter += new System.EventHandler(this.txtCuotaFacil_Enter);
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.Location = new System.Drawing.Point(16, 96);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(64, 16);
			this.label40.TabIndex = 13;
			this.label40.Text = "Otr Tarjetas";
			// 
			// txtOtrasTarjetas
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOtrasTarjetas.Appearance = appearance3;
			this.txtOtrasTarjetas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Tarjeta"));
			this.txtOtrasTarjetas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtrasTarjetas.Enabled = false;
			this.txtOtrasTarjetas.FormatString = "#,##0.00";
			this.txtOtrasTarjetas.Location = new System.Drawing.Point(80, 88);
			this.txtOtrasTarjetas.Name = "txtOtrasTarjetas";
			this.txtOtrasTarjetas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtOtrasTarjetas.PromptChar = ' ';
			this.txtOtrasTarjetas.Size = new System.Drawing.Size(104, 21);
			this.txtOtrasTarjetas.TabIndex = 6;
			this.txtOtrasTarjetas.Validated += new System.EventHandler(this.txtOtrasTarjetas_Validated);
			this.txtOtrasTarjetas.ValueChanged += new System.EventHandler(this.txtOtrasTarjetas_ValueChanged);
			// 
			// label41
			// 
			this.label41.AutoSize = true;
			this.label41.Location = new System.Drawing.Point(16, 120);
			this.label41.Name = "label41";
			this.label41.Size = new System.Drawing.Size(40, 16);
			this.label41.TabIndex = 15;
			this.label41.Text = "Gastos";
			// 
			// txtGastos
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGastos.Appearance = appearance4;
			this.txtGastos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Gastos"));
			this.txtGastos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastos.FormatString = "#,##0.00";
			this.txtGastos.Location = new System.Drawing.Point(80, 112);
			this.txtGastos.Name = "txtGastos";
			this.txtGastos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastos.PromptChar = ' ';
			this.txtGastos.Size = new System.Drawing.Size(104, 21);
			this.txtGastos.TabIndex = 8;
			this.txtGastos.Validated += new System.EventHandler(this.txtGastos_Validated);
			this.txtGastos.ValueChanged += new System.EventHandler(this.txtGastos_ValueChanged);
			this.txtGastos.Enter += new System.EventHandler(this.txtGastos_Enter);
			// 
			// label42
			// 
			this.label42.AutoSize = true;
			this.label42.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label42.Location = new System.Drawing.Point(192, 144);
			this.label42.Name = "label42";
			this.label42.Size = new System.Drawing.Size(64, 16);
			this.label42.TabIndex = 23;
			this.label42.Text = "DEPOSITO";
			// 
			// txtDeposito
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDeposito.Appearance = appearance5;
			this.txtDeposito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Deposito"));
			this.txtDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDeposito.Enabled = false;
			this.txtDeposito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDeposito.FormatString = "#,##0.00";
			this.txtDeposito.Location = new System.Drawing.Point(272, 136);
			this.txtDeposito.Name = "txtDeposito";
			this.txtDeposito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDeposito.PromptChar = ' ';
			this.txtDeposito.Size = new System.Drawing.Size(96, 21);
			this.txtDeposito.TabIndex = 11;
			// 
			// label43
			// 
			this.label43.AutoSize = true;
			this.label43.Location = new System.Drawing.Point(192, 96);
			this.label43.Name = "label43";
			this.label43.Size = new System.Drawing.Size(77, 16);
			this.label43.TabIndex = 21;
			this.label43.Text = "Cred Personal";
			// 
			// txtCredPersonal
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCredPersonal.Appearance = appearance6;
			this.txtCredPersonal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Credito"));
			this.txtCredPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCredPersonal.FormatString = "#,##0.00";
			this.txtCredPersonal.Location = new System.Drawing.Point(272, 88);
			this.txtCredPersonal.Name = "txtCredPersonal";
			this.txtCredPersonal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCredPersonal.PromptChar = ' ';
			this.txtCredPersonal.Size = new System.Drawing.Size(96, 21);
			this.txtCredPersonal.TabIndex = 7;
			this.txtCredPersonal.Validated += new System.EventHandler(this.txtCredPersonal_Validated);
			this.txtCredPersonal.ValueChanged += new System.EventHandler(this.txtCredPersonal_ValueChanged);
			this.txtCredPersonal.Enter += new System.EventHandler(this.txtCredPersonal_Enter);
			// 
			// label44
			// 
			this.label44.AutoSize = true;
			this.label44.Location = new System.Drawing.Point(192, 72);
			this.label44.Name = "label44";
			this.label44.Size = new System.Drawing.Size(73, 16);
			this.label44.TabIndex = 19;
			this.label44.Text = "Total Tarjetas";
			// 
			// txtTotalTarjetas
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalTarjetas.Appearance = appearance7;
			this.txtTotalTarjetas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalTarjetas.FormatString = "#,##0.00";
			this.txtTotalTarjetas.Location = new System.Drawing.Point(272, 64);
			this.txtTotalTarjetas.Name = "txtTotalTarjetas";
			this.txtTotalTarjetas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalTarjetas.PromptChar = ' ';
			this.txtTotalTarjetas.Size = new System.Drawing.Size(96, 21);
			this.txtTotalTarjetas.TabIndex = 5;
			this.txtTotalTarjetas.Validated += new System.EventHandler(this.txtTotalTarjetas_Validated);
			this.txtTotalTarjetas.ValueChanged += new System.EventHandler(this.txtTotalTarjetas_ValueChanged);
			this.txtTotalTarjetas.Enter += new System.EventHandler(this.txtTotalTarjetas_Enter);
			// 
			// label45
			// 
			this.label45.AutoSize = true;
			this.label45.Location = new System.Drawing.Point(192, 48);
			this.label45.Name = "label45";
			this.label45.Size = new System.Drawing.Size(50, 16);
			this.label45.TabIndex = 17;
			this.label45.Text = "Cheques";
			// 
			// txtCheque
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCheque.Appearance = appearance8;
			this.txtCheque.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Cheque"));
			this.txtCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCheque.FormatString = "#,##0.00";
			this.txtCheque.Location = new System.Drawing.Point(272, 40);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCheque.PromptChar = ' ';
			this.txtCheque.Size = new System.Drawing.Size(96, 21);
			this.txtCheque.TabIndex = 3;
			this.txtCheque.Validated += new System.EventHandler(this.txtCheque_Validated);
			this.txtCheque.ValueChanged += new System.EventHandler(this.txtCheque_ValueChanged);
			this.txtCheque.Enter += new System.EventHandler(this.txtCheque_Enter);
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label46.Location = new System.Drawing.Point(648, 64);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(54, 16);
			this.label46.TabIndex = 24;
			this.label46.Text = "VAL-ACT";
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label47.Location = new System.Drawing.Point(648, 88);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(52, 16);
			this.label47.TabIndex = 25;
			this.label47.Text = "UNI-ACT";
			// 
			// label48
			// 
			this.label48.AutoSize = true;
			this.label48.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label48.Location = new System.Drawing.Point(656, 112);
			this.label48.Name = "label48";
			this.label48.Size = new System.Drawing.Size(54, 16);
			this.label48.TabIndex = 26;
			this.label48.Text = "VAL-ANT";
			// 
			// label49
			// 
			this.label49.AutoSize = true;
			this.label49.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label49.Location = new System.Drawing.Point(664, 136);
			this.label49.Name = "label49";
			this.label49.Size = new System.Drawing.Size(52, 16);
			this.label49.TabIndex = 27;
			this.label49.Text = "UNI-ANT";
			// 
			// txtValorActual
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorActual.Appearance = appearance9;
			this.txtValorActual.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ValorActual"));
			this.txtValorActual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorActual.Enabled = false;
			this.txtValorActual.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtValorActual.FormatString = "#,##0.00";
			this.txtValorActual.Location = new System.Drawing.Point(712, 56);
			this.txtValorActual.Name = "txtValorActual";
			this.txtValorActual.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorActual.PromptChar = ' ';
			this.txtValorActual.Size = new System.Drawing.Size(72, 21);
			this.txtValorActual.TabIndex = 117;
			// 
			// txtUnidadActual
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUnidadActual.Appearance = appearance10;
			this.txtUnidadActual.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "UnidadesActual"));
			this.txtUnidadActual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUnidadActual.Enabled = false;
			this.txtUnidadActual.FormatString = "#,##0.00";
			this.txtUnidadActual.Location = new System.Drawing.Point(712, 80);
			this.txtUnidadActual.Name = "txtUnidadActual";
			this.txtUnidadActual.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUnidadActual.PromptChar = ' ';
			this.txtUnidadActual.Size = new System.Drawing.Size(72, 21);
			this.txtUnidadActual.TabIndex = 118;
			// 
			// txtValorAntes
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorAntes.Appearance = appearance11;
			this.txtValorAntes.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ValorAntes"));
			this.txtValorAntes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorAntes.Enabled = false;
			this.txtValorAntes.FormatString = "#,##0.00";
			this.txtValorAntes.Location = new System.Drawing.Point(728, 104);
			this.txtValorAntes.Name = "txtValorAntes";
			this.txtValorAntes.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorAntes.PromptChar = ' ';
			this.txtValorAntes.Size = new System.Drawing.Size(56, 21);
			this.txtValorAntes.TabIndex = 119;
			// 
			// txtUnidadesAntes
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUnidadesAntes.Appearance = appearance12;
			this.txtUnidadesAntes.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "UnidadesAntes"));
			this.txtUnidadesAntes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUnidadesAntes.Enabled = false;
			this.txtUnidadesAntes.FormatString = "#,##0.00";
			this.txtUnidadesAntes.Location = new System.Drawing.Point(728, 128);
			this.txtUnidadesAntes.Name = "txtUnidadesAntes";
			this.txtUnidadesAntes.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUnidadesAntes.PromptChar = ' ';
			this.txtUnidadesAntes.Size = new System.Drawing.Size(56, 21);
			this.txtUnidadesAntes.TabIndex = 120;
			// 
			// label50
			// 
			this.label50.AutoSize = true;
			this.label50.Location = new System.Drawing.Point(792, 64);
			this.label50.Name = "label50";
			this.label50.Size = new System.Drawing.Size(44, 16);
			this.label50.TabIndex = 32;
			this.label50.Text = "Hombre";
			// 
			// label51
			// 
			this.label51.AutoSize = true;
			this.label51.Location = new System.Drawing.Point(800, 88);
			this.label51.Name = "label51";
			this.label51.Size = new System.Drawing.Size(33, 16);
			this.label51.TabIndex = 33;
			this.label51.Text = "Mujer";
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.Location = new System.Drawing.Point(928, 312);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(32, 16);
			this.label53.TabIndex = 35;
			this.label53.Text = "Otros";
			// 
			// label54
			// 
			this.label54.AutoSize = true;
			this.label54.Location = new System.Drawing.Point(912, 232);
			this.label54.Name = "label54";
			this.label54.Size = new System.Drawing.Size(53, 16);
			this.label54.TabIndex = 36;
			this.label54.Text = "Promo 29";
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.Location = new System.Drawing.Point(912, 160);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(53, 16);
			this.label55.TabIndex = 37;
			this.label55.Text = "Promo 15";
			// 
			// label56
			// 
			this.label56.AutoSize = true;
			this.label56.Location = new System.Drawing.Point(912, 136);
			this.label56.Name = "label56";
			this.label56.Size = new System.Drawing.Size(53, 16);
			this.label56.TabIndex = 38;
			this.label56.Text = "Promo 10";
			// 
			// label57
			// 
			this.label57.AutoSize = true;
			this.label57.Location = new System.Drawing.Point(920, 112);
			this.label57.Name = "label57";
			this.label57.Size = new System.Drawing.Size(47, 16);
			this.label57.TabIndex = 39;
			this.label57.Text = "Promo 5";
			// 
			// label58
			// 
			this.label58.AutoSize = true;
			this.label58.Location = new System.Drawing.Point(912, 88);
			this.label58.Name = "label58";
			this.label58.Size = new System.Drawing.Size(45, 16);
			this.label58.TabIndex = 40;
			this.label58.Text = "Balones";
			// 
			// label59
			// 
			this.label59.AutoSize = true;
			this.label59.Location = new System.Drawing.Point(888, 64);
			this.label59.Name = "label59";
			this.label59.Size = new System.Drawing.Size(71, 16);
			this.label59.TabIndex = 41;
			this.label59.Text = "Portazapatos";
			// 
			// label60
			// 
			this.label60.AutoSize = true;
			this.label60.Location = new System.Drawing.Point(792, 304);
			this.label60.Name = "label60";
			this.label60.Size = new System.Drawing.Size(45, 16);
			this.label60.TabIndex = 42;
			this.label60.Text = "Billetera";
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.Location = new System.Drawing.Point(784, 280);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(48, 16);
			this.label61.TabIndex = 43;
			this.label61.Text = "Canguro";
			// 
			// label62
			// 
			this.label62.AutoSize = true;
			this.label62.Location = new System.Drawing.Point(800, 256);
			this.label62.Name = "label62";
			this.label62.Size = new System.Drawing.Size(38, 16);
			this.label62.TabIndex = 44;
			this.label62.Text = "Maleta";
			// 
			// label63
			// 
			this.label63.AutoSize = true;
			this.label63.Location = new System.Drawing.Point(792, 232);
			this.label63.Name = "label63";
			this.label63.Size = new System.Drawing.Size(43, 16);
			this.label63.TabIndex = 45;
			this.label63.Text = "Mochila";
			// 
			// label64
			// 
			this.label64.AutoSize = true;
			this.label64.Location = new System.Drawing.Point(800, 208);
			this.label64.Name = "label64";
			this.label64.Size = new System.Drawing.Size(33, 16);
			this.label64.TabIndex = 46;
			this.label64.Text = "Gorra";
			// 
			// label65
			// 
			this.label65.AutoSize = true;
			this.label65.Location = new System.Drawing.Point(776, 184);
			this.label65.Name = "label65";
			this.label65.Size = new System.Drawing.Size(59, 16);
			this.label65.TabIndex = 47;
			this.label65.Text = "Cinturones";
			// 
			// label66
			// 
			this.label66.AutoSize = true;
			this.label66.Location = new System.Drawing.Point(784, 160);
			this.label66.Name = "label66";
			this.label66.Size = new System.Drawing.Size(48, 16);
			this.label66.TabIndex = 48;
			this.label66.Text = "Carteras";
			// 
			// label67
			// 
			this.label67.AutoSize = true;
			this.label67.Location = new System.Drawing.Point(808, 136);
			this.label67.Name = "label67";
			this.label67.Size = new System.Drawing.Size(28, 16);
			this.label67.TabIndex = 49;
			this.label67.Text = "Niña";
			// 
			// label68
			// 
			this.label68.AutoSize = true;
			this.label68.Location = new System.Drawing.Point(808, 112);
			this.label68.Name = "label68";
			this.label68.Size = new System.Drawing.Size(28, 16);
			this.label68.TabIndex = 50;
			this.label68.Text = "Niño";
			// 
			// txtHombre
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtHombre.Appearance = appearance13;
			this.txtHombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Hombre"));
			this.txtHombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHombre.FormatString = "0";
			this.txtHombre.Location = new System.Drawing.Point(848, 64);
			this.txtHombre.MaskInput = "nnn";
			this.txtHombre.Name = "txtHombre";
			this.txtHombre.PromptChar = ' ';
			this.txtHombre.Size = new System.Drawing.Size(32, 21);
			this.txtHombre.TabIndex = 121;
			this.txtHombre.Validated += new System.EventHandler(this.txtHombre_Validated);
			this.txtHombre.Enter += new System.EventHandler(this.txtHombre_Enter);
			// 
			// txtMujer
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMujer.Appearance = appearance14;
			this.txtMujer.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Mujer"));
			this.txtMujer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMujer.FormatString = "0";
			this.txtMujer.Location = new System.Drawing.Point(848, 88);
			this.txtMujer.MaskInput = "nnn";
			this.txtMujer.Name = "txtMujer";
			this.txtMujer.PromptChar = ' ';
			this.txtMujer.Size = new System.Drawing.Size(32, 21);
			this.txtMujer.TabIndex = 122;
			this.txtMujer.Validated += new System.EventHandler(this.txtMujer_Validated);
			this.txtMujer.Enter += new System.EventHandler(this.txtMujer_Enter);
			// 
			// txtNiño
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNiño.Appearance = appearance15;
			this.txtNiño.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Niño"));
			this.txtNiño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNiño.FormatString = "0";
			this.txtNiño.Location = new System.Drawing.Point(848, 112);
			this.txtNiño.MaskInput = "nnn";
			this.txtNiño.Name = "txtNiño";
			this.txtNiño.PromptChar = ' ';
			this.txtNiño.Size = new System.Drawing.Size(32, 21);
			this.txtNiño.TabIndex = 123;
			this.txtNiño.Validated += new System.EventHandler(this.txtNiño_Validated);
			this.txtNiño.Enter += new System.EventHandler(this.txtNiño_Enter);
			// 
			// lblfecha
			// 
			this.lblfecha.Location = new System.Drawing.Point(384, 40);
			this.lblfecha.Name = "lblfecha";
			this.lblfecha.Size = new System.Drawing.Size(88, 23);
			this.lblfecha.TabIndex = 134;
			this.lblfecha.Text = "label69";
			// 
			// cdsMarceloBodega
			// 
			this.cdsMarceloBodega.BindingContextCtrl = this;
			this.cdsMarceloBodega.DataLibrary = "LibFacturacion";
			this.cdsMarceloBodega.DataLibraryUrl = "";
			this.cdsMarceloBodega.DataSetDef = "dsMarceloBodega";
			this.cdsMarceloBodega.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsMarceloBodega.Name = "cdsMarceloBodega";
			this.cdsMarceloBodega.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMarceloBodega.SchemaDef = null;
			this.cdsMarceloBodega.BeforeFill += new C1.Data.FillEventHandler(this.cdsMarceloBodega_BeforeFill);
			// 
			// cdvMarceloBodega
			// 
			this.cdvMarceloBodega.BindingContextCtrl = this;
			this.cdvMarceloBodega.DataSet = this.cdsMarceloBodega;
			this.cdvMarceloBodega.RowFilter = "status=1";
			this.cdvMarceloBodega.Sort = "bodega";
			this.cdvMarceloBodega.TableName = "";
			this.cdvMarceloBodega.TableViewName = "MarceloBodega";
			// 
			// cmbBodega
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance16;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.cdvMarceloBodega;
			this.cmbBodega.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "MarceloBodega";
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Bodega";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(80, 8);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(288, 21);
			this.cmbBodega.TabIndex = 0;
			this.cmbBodega.ValueMember = "Clave";
			this.cmbBodega.RowSelected += new Infragistics.Win.UltraWinGrid.RowSelectedEventHandler(this.cmbBodega_RowSelected_1);
			// 
			// dtFecha
			// 
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MMM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(392, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(120, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1;
			this.dtFecha.Value = new System.DateTime(2008, 2, 21, 0, 0, 0, 0);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtPromo29
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPromo29.Appearance = appearance17;
			this.txtPromo29.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Promocion29"));
			this.txtPromo29.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPromo29.FormatString = "0";
			this.txtPromo29.Location = new System.Drawing.Point(968, 232);
			this.txtPromo29.MaskInput = "nnn";
			this.txtPromo29.Name = "txtPromo29";
			this.txtPromo29.PromptChar = ' ';
			this.txtPromo29.Size = new System.Drawing.Size(32, 21);
			this.txtPromo29.TabIndex = 139;
			this.txtPromo29.Validated += new System.EventHandler(this.txtPromo29_Validated);
			this.txtPromo29.Enter += new System.EventHandler(this.txtPromo29_Enter);
			// 
			// txtPromo15
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPromo15.Appearance = appearance18;
			this.txtPromo15.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Promocion15"));
			this.txtPromo15.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPromo15.FormatString = "0";
			this.txtPromo15.Location = new System.Drawing.Point(968, 160);
			this.txtPromo15.MaskInput = "nnn";
			this.txtPromo15.Name = "txtPromo15";
			this.txtPromo15.PromptChar = ' ';
			this.txtPromo15.Size = new System.Drawing.Size(32, 21);
			this.txtPromo15.TabIndex = 136;
			this.txtPromo15.Validated += new System.EventHandler(this.txtPromo15_Validated);
			this.txtPromo15.Enter += new System.EventHandler(this.txtPromo15_Enter);
			// 
			// txtPromo10
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPromo10.Appearance = appearance19;
			this.txtPromo10.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Promocion10"));
			this.txtPromo10.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPromo10.FormatString = "0";
			this.txtPromo10.Location = new System.Drawing.Point(968, 136);
			this.txtPromo10.MaskInput = "nnn";
			this.txtPromo10.Name = "txtPromo10";
			this.txtPromo10.PromptChar = ' ';
			this.txtPromo10.Size = new System.Drawing.Size(32, 21);
			this.txtPromo10.TabIndex = 135;
			this.txtPromo10.Validated += new System.EventHandler(this.txtPromo10_Validated);
			this.txtPromo10.Enter += new System.EventHandler(this.txtPromo10_Enter);
			// 
			// txtPromo5
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPromo5.Appearance = appearance20;
			this.txtPromo5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Promocion5"));
			this.txtPromo5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPromo5.FormatString = "0";
			this.txtPromo5.Location = new System.Drawing.Point(968, 112);
			this.txtPromo5.MaskInput = "nnn";
			this.txtPromo5.Name = "txtPromo5";
			this.txtPromo5.PromptChar = ' ';
			this.txtPromo5.Size = new System.Drawing.Size(32, 21);
			this.txtPromo5.TabIndex = 134;
			this.txtPromo5.Validated += new System.EventHandler(this.txtPromo5_Validated);
			this.txtPromo5.Enter += new System.EventHandler(this.txtPromo5_Enter);
			// 
			// txtBalon
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBalon.Appearance = appearance21;
			this.txtBalon.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Bolso"));
			this.txtBalon.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBalon.FormatString = "0";
			this.txtBalon.Location = new System.Drawing.Point(968, 88);
			this.txtBalon.MaskInput = "nnn";
			this.txtBalon.Name = "txtBalon";
			this.txtBalon.PromptChar = ' ';
			this.txtBalon.Size = new System.Drawing.Size(32, 21);
			this.txtBalon.TabIndex = 133;
			this.txtBalon.Validated += new System.EventHandler(this.txtBalon_Validated);
			this.txtBalon.Enter += new System.EventHandler(this.txtBalon_Enter);
			// 
			// txtPortazapatos
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPortazapatos.Appearance = appearance22;
			this.txtPortazapatos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "PortaZapatos"));
			this.txtPortazapatos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPortazapatos.FormatString = "0";
			this.txtPortazapatos.Location = new System.Drawing.Point(968, 64);
			this.txtPortazapatos.MaskInput = "nnn";
			this.txtPortazapatos.Name = "txtPortazapatos";
			this.txtPortazapatos.PromptChar = ' ';
			this.txtPortazapatos.Size = new System.Drawing.Size(32, 21);
			this.txtPortazapatos.TabIndex = 132;
			this.txtPortazapatos.Validated += new System.EventHandler(this.txtPortazapatos_Validated);
			this.txtPortazapatos.Enter += new System.EventHandler(this.txtPortazapatos_Enter);
			// 
			// txtBilletera
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBilletera.Appearance = appearance23;
			this.txtBilletera.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Billetera"));
			this.txtBilletera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBilletera.FormatString = "0";
			this.txtBilletera.Location = new System.Drawing.Point(848, 304);
			this.txtBilletera.MaskInput = "nnn";
			this.txtBilletera.Name = "txtBilletera";
			this.txtBilletera.PromptChar = ' ';
			this.txtBilletera.Size = new System.Drawing.Size(32, 21);
			this.txtBilletera.TabIndex = 131;
			this.txtBilletera.Validated += new System.EventHandler(this.txtBilletera_Validated);
			this.txtBilletera.Enter += new System.EventHandler(this.txtBilletera_Enter);
			// 
			// txtCanguro
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCanguro.Appearance = appearance24;
			this.txtCanguro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Canguro"));
			this.txtCanguro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCanguro.FormatString = "0";
			this.txtCanguro.Location = new System.Drawing.Point(848, 280);
			this.txtCanguro.MaskInput = "nnn";
			this.txtCanguro.Name = "txtCanguro";
			this.txtCanguro.PromptChar = ' ';
			this.txtCanguro.Size = new System.Drawing.Size(32, 21);
			this.txtCanguro.TabIndex = 130;
			this.txtCanguro.Validated += new System.EventHandler(this.txtCanguro_Validated);
			this.txtCanguro.Enter += new System.EventHandler(this.txtCanguro_Enter);
			// 
			// txtMaleta
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMaleta.Appearance = appearance25;
			this.txtMaleta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Maleta"));
			this.txtMaleta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMaleta.FormatString = "0";
			this.txtMaleta.Location = new System.Drawing.Point(848, 256);
			this.txtMaleta.MaskInput = "nnn";
			this.txtMaleta.Name = "txtMaleta";
			this.txtMaleta.PromptChar = ' ';
			this.txtMaleta.Size = new System.Drawing.Size(32, 21);
			this.txtMaleta.TabIndex = 129;
			this.txtMaleta.Validated += new System.EventHandler(this.txtMaleta_Validated);
			this.txtMaleta.Enter += new System.EventHandler(this.txtMaleta_Enter);
			// 
			// txtMochila
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMochila.Appearance = appearance26;
			this.txtMochila.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Mochila"));
			this.txtMochila.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMochila.FormatString = "0";
			this.txtMochila.Location = new System.Drawing.Point(848, 232);
			this.txtMochila.MaskInput = "nnn";
			this.txtMochila.Name = "txtMochila";
			this.txtMochila.PromptChar = ' ';
			this.txtMochila.Size = new System.Drawing.Size(32, 21);
			this.txtMochila.TabIndex = 128;
			this.txtMochila.Validated += new System.EventHandler(this.txtMochila_Validated);
			this.txtMochila.Enter += new System.EventHandler(this.txtMochila_Enter);
			// 
			// txtGorra
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGorra.Appearance = appearance27;
			this.txtGorra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Gorra"));
			this.txtGorra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGorra.FormatString = "0";
			this.txtGorra.Location = new System.Drawing.Point(848, 208);
			this.txtGorra.MaskInput = "nnn";
			this.txtGorra.Name = "txtGorra";
			this.txtGorra.PromptChar = ' ';
			this.txtGorra.Size = new System.Drawing.Size(32, 21);
			this.txtGorra.TabIndex = 127;
			this.txtGorra.Validated += new System.EventHandler(this.txtGorra_Validated);
			this.txtGorra.Enter += new System.EventHandler(this.txtGorra_Enter);
			// 
			// txtCinturon
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCinturon.Appearance = appearance28;
			this.txtCinturon.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Cinturon"));
			this.txtCinturon.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCinturon.FormatString = "0";
			this.txtCinturon.Location = new System.Drawing.Point(848, 184);
			this.txtCinturon.MaskInput = "nnn";
			this.txtCinturon.Name = "txtCinturon";
			this.txtCinturon.PromptChar = ' ';
			this.txtCinturon.Size = new System.Drawing.Size(32, 21);
			this.txtCinturon.TabIndex = 126;
			this.txtCinturon.Validated += new System.EventHandler(this.txtCinturon_Validated);
			this.txtCinturon.Enter += new System.EventHandler(this.txtCinturon_Enter);
			// 
			// txtCartera
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCartera.Appearance = appearance29;
			this.txtCartera.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Cartera"));
			this.txtCartera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCartera.FormatString = "0";
			this.txtCartera.Location = new System.Drawing.Point(848, 160);
			this.txtCartera.MaskInput = "nnn";
			this.txtCartera.Name = "txtCartera";
			this.txtCartera.PromptChar = ' ';
			this.txtCartera.Size = new System.Drawing.Size(32, 21);
			this.txtCartera.TabIndex = 125;
			this.txtCartera.Validated += new System.EventHandler(this.txtCartera_Validated);
			this.txtCartera.Enter += new System.EventHandler(this.txtCartera_Enter);
			// 
			// txtNiña
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNiña.Appearance = appearance30;
			this.txtNiña.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Niña"));
			this.txtNiña.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNiña.FormatString = "0";
			this.txtNiña.Location = new System.Drawing.Point(848, 136);
			this.txtNiña.MaskInput = "nnn";
			this.txtNiña.Name = "txtNiña";
			this.txtNiña.PromptChar = ' ';
			this.txtNiña.Size = new System.Drawing.Size(32, 21);
			this.txtNiña.TabIndex = 124;
			this.txtNiña.Validated += new System.EventHandler(this.txtNiña_Validated);
			this.txtNiña.Enter += new System.EventHandler(this.txtNiña_Enter);
			// 
			// txtTotal
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance31;
			this.txtTotal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "UnidadesActual"));
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.FormatString = "0";
			this.txtTotal.Location = new System.Drawing.Point(912, 328);
			this.txtTotal.MaskInput = "nnnnnnnnn";
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(48, 21);
			this.txtTotal.TabIndex = 143;
			// 
			// txtPromoOtro
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPromoOtro.Appearance = appearance32;
			this.txtPromoOtro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Otro"));
			this.txtPromoOtro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPromoOtro.FormatString = "0";
			this.txtPromoOtro.Location = new System.Drawing.Point(968, 304);
			this.txtPromoOtro.MaskInput = "nnn";
			this.txtPromoOtro.Name = "txtPromoOtro";
			this.txtPromoOtro.PromptChar = ' ';
			this.txtPromoOtro.Size = new System.Drawing.Size(32, 21);
			this.txtPromoOtro.TabIndex = 142;
			this.txtPromoOtro.Validated += new System.EventHandler(this.txtPromoOtro_Validated);
			this.txtPromoOtro.Enter += new System.EventHandler(this.txtPromoOtro_Enter);
			// 
			// label69
			// 
			this.label69.AutoSize = true;
			this.label69.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label69.Location = new System.Drawing.Point(864, 328);
			this.label69.Name = "label69";
			this.label69.Size = new System.Drawing.Size(42, 16);
			this.label69.TabIndex = 154;
			this.label69.Text = "TOTAL";
			// 
			// label52
			// 
			this.label52.AutoSize = true;
			this.label52.Location = new System.Drawing.Point(912, 184);
			this.label52.Name = "label52";
			this.label52.Size = new System.Drawing.Size(53, 16);
			this.label52.TabIndex = 155;
			this.label52.Text = "Promo 20";
			// 
			// txtPromo20
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPromo20.Appearance = appearance33;
			this.txtPromo20.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Promocion20"));
			this.txtPromo20.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPromo20.FormatString = "0";
			this.txtPromo20.Location = new System.Drawing.Point(968, 184);
			this.txtPromo20.MaskInput = "nnn";
			this.txtPromo20.Name = "txtPromo20";
			this.txtPromo20.PromptChar = ' ';
			this.txtPromo20.Size = new System.Drawing.Size(32, 21);
			this.txtPromo20.TabIndex = 137;
			this.txtPromo20.Validated += new System.EventHandler(this.txtPromo20_Validated);
			this.txtPromo20.Enter += new System.EventHandler(this.txtPromo20_Enter);
			// 
			// txtPromo25
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPromo25.Appearance = appearance34;
			this.txtPromo25.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Promocion25"));
			this.txtPromo25.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPromo25.FormatString = "0";
			this.txtPromo25.Location = new System.Drawing.Point(968, 208);
			this.txtPromo25.MaskInput = "nnn";
			this.txtPromo25.Name = "txtPromo25";
			this.txtPromo25.PromptChar = ' ';
			this.txtPromo25.Size = new System.Drawing.Size(32, 21);
			this.txtPromo25.TabIndex = 138;
			this.txtPromo25.Validated += new System.EventHandler(this.txtPromo25_Validated);
			this.txtPromo25.Enter += new System.EventHandler(this.txtPromo25_Enter);
			// 
			// label70
			// 
			this.label70.AutoSize = true;
			this.label70.Location = new System.Drawing.Point(912, 208);
			this.label70.Name = "label70";
			this.label70.Size = new System.Drawing.Size(53, 16);
			this.label70.TabIndex = 157;
			this.label70.Text = "Promo 25";
			// 
			// btGrabar
			// 
			this.btGrabar.Location = new System.Drawing.Point(680, 552);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.TabIndex = 148;
			this.btGrabar.Text = "Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btDeshacer
			// 
			this.btDeshacer.Location = new System.Drawing.Point(760, 552);
			this.btDeshacer.Name = "btDeshacer";
			this.btDeshacer.TabIndex = 149;
			this.btDeshacer.Text = "Deshacer";
			this.btDeshacer.Click += new System.EventHandler(this.btDeshacer_Click);
			// 
			// btReportes
			// 
			this.btReportes.Location = new System.Drawing.Point(928, 552);
			this.btReportes.Name = "btReportes";
			this.btReportes.TabIndex = 150;
			this.btReportes.Text = "Reportes";
			this.btReportes.Click += new System.EventHandler(this.btReportes_Click);
			// 
			// label71
			// 
			this.label71.AutoSize = true;
			this.label71.Location = new System.Drawing.Point(8, 344);
			this.label71.Name = "label71";
			this.label71.Size = new System.Drawing.Size(64, 16);
			this.label71.TabIndex = 161;
			this.label71.Text = "Cuota Facil ";
			this.label71.Click += new System.EventHandler(this.label71_Click);
			// 
			// label73
			// 
			this.label73.AutoSize = true;
			this.label73.Location = new System.Drawing.Point(8, 392);
			this.label73.Name = "label73";
			this.label73.Size = new System.Drawing.Size(87, 16);
			this.label73.TabIndex = 163;
			this.label73.Text = "Diners Corriente";
			// 
			// label74
			// 
			this.label74.AutoSize = true;
			this.label74.Location = new System.Drawing.Point(8, 416);
			this.label74.Name = "label74";
			this.label74.Size = new System.Drawing.Size(79, 16);
			this.label74.TabIndex = 164;
			this.label74.Text = "Diners Diferido";
			// 
			// label75
			// 
			this.label75.AutoSize = true;
			this.label75.Location = new System.Drawing.Point(8, 440);
			this.label75.Name = "label75";
			this.label75.Size = new System.Drawing.Size(54, 16);
			this.label75.TabIndex = 165;
			this.label75.Text = "Pacificard";
			// 
			// label76
			// 
			this.label76.AutoSize = true;
			this.label76.Location = new System.Drawing.Point(8, 464);
			this.label76.Name = "label76";
			this.label76.Size = new System.Drawing.Size(79, 16);
			this.label76.TabIndex = 166;
			this.label76.Text = "Pichincha Visa";
			// 
			// label77
			// 
			this.label77.AutoSize = true;
			this.label77.Location = new System.Drawing.Point(8, 368);
			this.label77.Name = "label77";
			this.label77.Size = new System.Drawing.Size(113, 16);
			this.label77.TabIndex = 167;
			this.label77.Text = "Pichincha Mastercard";
			// 
			// label78
			// 
			this.label78.AutoSize = true;
			this.label78.Location = new System.Drawing.Point(8, 488);
			this.label78.Name = "label78";
			this.label78.Size = new System.Drawing.Size(96, 16);
			this.label78.TabIndex = 168;
			this.label78.Text = "American Express";
			// 
			// label79
			// 
			this.label79.AutoSize = true;
			this.label79.Location = new System.Drawing.Point(8, 512);
			this.label79.Name = "label79";
			this.label79.Size = new System.Drawing.Size(52, 16);
			this.label79.TabIndex = 169;
			this.label79.Text = "Territorial";
			// 
			// label80
			// 
			this.label80.AutoSize = true;
			this.label80.Location = new System.Drawing.Point(8, 536);
			this.label80.Name = "label80";
			this.label80.Size = new System.Drawing.Size(60, 16);
			this.label80.TabIndex = 170;
			this.label80.Text = "Bolivariano";
			this.label80.Click += new System.EventHandler(this.label80_Click);
			// 
			// label81
			// 
			this.label81.AutoSize = true;
			this.label81.Location = new System.Drawing.Point(16, 560);
			this.label81.Name = "label81";
			this.label81.Size = new System.Drawing.Size(37, 16);
			this.label81.TabIndex = 171;
			this.label81.Text = "Austro";
			// 
			// ultraNumericEditor1
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor1.Appearance = appearance35;
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "c8"));
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.FormatString = "";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(128, 336);
			this.ultraNumericEditor1.MaskInput = "nnn";
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor1.TabIndex = 27;
			this.ultraNumericEditor1.Enter += new System.EventHandler(this.ultraNumericEditor1_Enter);
			// 
			// ultraNumericEditor3
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor3.Appearance = appearance36;
			this.ultraNumericEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "c1"));
			this.ultraNumericEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor3.FormatString = "";
			this.ultraNumericEditor3.Location = new System.Drawing.Point(128, 384);
			this.ultraNumericEditor3.MaskInput = "nnn";
			this.ultraNumericEditor3.Name = "ultraNumericEditor3";
			this.ultraNumericEditor3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor3.PromptChar = ' ';
			this.ultraNumericEditor3.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor3.TabIndex = 33;
			this.ultraNumericEditor3.Enter += new System.EventHandler(this.ultraNumericEditor3_Enter);
			// 
			// ultraNumericEditor4
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor4.Appearance = appearance37;
			this.ultraNumericEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "c2"));
			this.ultraNumericEditor4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor4.FormatString = "";
			this.ultraNumericEditor4.Location = new System.Drawing.Point(128, 408);
			this.ultraNumericEditor4.MaskInput = "nnn";
			this.ultraNumericEditor4.Name = "ultraNumericEditor4";
			this.ultraNumericEditor4.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor4.PromptChar = ' ';
			this.ultraNumericEditor4.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor4.TabIndex = 36;
			this.ultraNumericEditor4.Enter += new System.EventHandler(this.ultraNumericEditor4_Enter);
			// 
			// ultraNumericEditor5
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor5.Appearance = appearance38;
			this.ultraNumericEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "c3"));
			this.ultraNumericEditor5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor5.FormatString = "";
			this.ultraNumericEditor5.Location = new System.Drawing.Point(128, 432);
			this.ultraNumericEditor5.MaskInput = "nnn";
			this.ultraNumericEditor5.Name = "ultraNumericEditor5";
			this.ultraNumericEditor5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor5.PromptChar = ' ';
			this.ultraNumericEditor5.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor5.TabIndex = 39;
			this.ultraNumericEditor5.Enter += new System.EventHandler(this.ultraNumericEditor5_Enter);
			// 
			// ultraNumericEditor6
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor6.Appearance = appearance39;
			this.ultraNumericEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "c4"));
			this.ultraNumericEditor6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor6.FormatString = "";
			this.ultraNumericEditor6.Location = new System.Drawing.Point(128, 456);
			this.ultraNumericEditor6.MaskInput = "nnn";
			this.ultraNumericEditor6.Name = "ultraNumericEditor6";
			this.ultraNumericEditor6.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor6.PromptChar = ' ';
			this.ultraNumericEditor6.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor6.TabIndex = 42;
			this.ultraNumericEditor6.Enter += new System.EventHandler(this.ultraNumericEditor6_Enter);
			// 
			// ultraNumericEditor7
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor7.Appearance = appearance40;
			this.ultraNumericEditor7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "c7"));
			this.ultraNumericEditor7.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor7.FormatString = "";
			this.ultraNumericEditor7.Location = new System.Drawing.Point(128, 360);
			this.ultraNumericEditor7.MaskInput = "nnn";
			this.ultraNumericEditor7.Name = "ultraNumericEditor7";
			this.ultraNumericEditor7.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor7.PromptChar = ' ';
			this.ultraNumericEditor7.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor7.TabIndex = 30;
			this.ultraNumericEditor7.Enter += new System.EventHandler(this.ultraNumericEditor7_Enter);
			// 
			// ultraNumericEditor8
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor8.Appearance = appearance41;
			this.ultraNumericEditor8.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "c5"));
			this.ultraNumericEditor8.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor8.FormatString = "";
			this.ultraNumericEditor8.Location = new System.Drawing.Point(128, 480);
			this.ultraNumericEditor8.MaskInput = "nnn";
			this.ultraNumericEditor8.Name = "ultraNumericEditor8";
			this.ultraNumericEditor8.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor8.PromptChar = ' ';
			this.ultraNumericEditor8.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor8.TabIndex = 45;
			this.ultraNumericEditor8.ValueChanged += new System.EventHandler(this.ultraNumericEditor8_ValueChanged);
			this.ultraNumericEditor8.Enter += new System.EventHandler(this.ultraNumericEditor8_Enter);
			// 
			// ultraNumericEditor9
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor9.Appearance = appearance42;
			this.ultraNumericEditor9.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "c6"));
			this.ultraNumericEditor9.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor9.FormatString = "";
			this.ultraNumericEditor9.Location = new System.Drawing.Point(128, 504);
			this.ultraNumericEditor9.MaskInput = "nnn";
			this.ultraNumericEditor9.Name = "ultraNumericEditor9";
			this.ultraNumericEditor9.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor9.PromptChar = ' ';
			this.ultraNumericEditor9.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor9.TabIndex = 48;
			this.ultraNumericEditor9.Enter += new System.EventHandler(this.ultraNumericEditor9_Enter);
			// 
			// ultraNumericEditor10
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor10.Appearance = appearance43;
			this.ultraNumericEditor10.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "c9"));
			this.ultraNumericEditor10.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor10.FormatString = "";
			this.ultraNumericEditor10.Location = new System.Drawing.Point(128, 528);
			this.ultraNumericEditor10.MaskInput = "nnn";
			this.ultraNumericEditor10.Name = "ultraNumericEditor10";
			this.ultraNumericEditor10.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor10.PromptChar = ' ';
			this.ultraNumericEditor10.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor10.TabIndex = 51;
			this.ultraNumericEditor10.Enter += new System.EventHandler(this.ultraNumericEditor10_Enter);
			// 
			// ultraNumericEditor11
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor11.Appearance = appearance44;
			this.ultraNumericEditor11.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "c10"));
			this.ultraNumericEditor11.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor11.FormatString = "";
			this.ultraNumericEditor11.Location = new System.Drawing.Point(128, 552);
			this.ultraNumericEditor11.MaskInput = "nnn";
			this.ultraNumericEditor11.Name = "ultraNumericEditor11";
			this.ultraNumericEditor11.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor11.PromptChar = ' ';
			this.ultraNumericEditor11.Size = new System.Drawing.Size(48, 21);
			this.ultraNumericEditor11.TabIndex = 54;
			this.ultraNumericEditor11.Enter += new System.EventHandler(this.ultraNumericEditor11_Enter);
			// 
			// ultraNumericEditor12
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor12.Appearance = appearance45;
			this.ultraNumericEditor12.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "l10"));
			this.ultraNumericEditor12.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor12.FormatString = "";
			this.ultraNumericEditor12.Location = new System.Drawing.Point(184, 552);
			this.ultraNumericEditor12.MaskInput = "nnnnnnn";
			this.ultraNumericEditor12.Name = "ultraNumericEditor12";
			this.ultraNumericEditor12.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor12.PromptChar = ' ';
			this.ultraNumericEditor12.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor12.TabIndex = 55;
			this.ultraNumericEditor12.Enter += new System.EventHandler(this.ultraNumericEditor12_Enter);
			// 
			// ultraNumericEditor13
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor13.Appearance = appearance46;
			this.ultraNumericEditor13.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "l9"));
			this.ultraNumericEditor13.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor13.FormatString = "";
			this.ultraNumericEditor13.Location = new System.Drawing.Point(184, 528);
			this.ultraNumericEditor13.MaskInput = "nnnnnnn";
			this.ultraNumericEditor13.Name = "ultraNumericEditor13";
			this.ultraNumericEditor13.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor13.PromptChar = ' ';
			this.ultraNumericEditor13.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor13.TabIndex = 52;
			this.ultraNumericEditor13.Enter += new System.EventHandler(this.ultraNumericEditor13_Enter);
			// 
			// ultraNumericEditor14
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor14.Appearance = appearance47;
			this.ultraNumericEditor14.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "l6"));
			this.ultraNumericEditor14.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor14.FormatString = "";
			this.ultraNumericEditor14.Location = new System.Drawing.Point(184, 504);
			this.ultraNumericEditor14.MaskInput = "nnnnnnn";
			this.ultraNumericEditor14.Name = "ultraNumericEditor14";
			this.ultraNumericEditor14.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor14.PromptChar = ' ';
			this.ultraNumericEditor14.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor14.TabIndex = 49;
			this.ultraNumericEditor14.Enter += new System.EventHandler(this.ultraNumericEditor14_Enter);
			// 
			// ultraNumericEditor15
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor15.Appearance = appearance48;
			this.ultraNumericEditor15.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "l5"));
			this.ultraNumericEditor15.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor15.FormatString = "";
			this.ultraNumericEditor15.Location = new System.Drawing.Point(184, 480);
			this.ultraNumericEditor15.MaskInput = "nnnnnnn";
			this.ultraNumericEditor15.Name = "ultraNumericEditor15";
			this.ultraNumericEditor15.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor15.PromptChar = ' ';
			this.ultraNumericEditor15.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor15.TabIndex = 46;
			this.ultraNumericEditor15.Enter += new System.EventHandler(this.ultraNumericEditor15_Enter);
			// 
			// ultraNumericEditor16
			// 
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor16.Appearance = appearance49;
			this.ultraNumericEditor16.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "l7"));
			this.ultraNumericEditor16.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor16.FormatString = "";
			this.ultraNumericEditor16.Location = new System.Drawing.Point(184, 360);
			this.ultraNumericEditor16.MaskInput = "nnnnnnn";
			this.ultraNumericEditor16.Name = "ultraNumericEditor16";
			this.ultraNumericEditor16.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor16.PromptChar = ' ';
			this.ultraNumericEditor16.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor16.TabIndex = 31;
			this.ultraNumericEditor16.Enter += new System.EventHandler(this.ultraNumericEditor16_Enter);
			// 
			// ultraNumericEditor17
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor17.Appearance = appearance50;
			this.ultraNumericEditor17.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "l4"));
			this.ultraNumericEditor17.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor17.FormatString = "";
			this.ultraNumericEditor17.Location = new System.Drawing.Point(184, 456);
			this.ultraNumericEditor17.MaskInput = "nnnnnnn";
			this.ultraNumericEditor17.Name = "ultraNumericEditor17";
			this.ultraNumericEditor17.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor17.PromptChar = ' ';
			this.ultraNumericEditor17.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor17.TabIndex = 43;
			this.ultraNumericEditor17.Enter += new System.EventHandler(this.ultraNumericEditor17_Enter);
			// 
			// ultraNumericEditor18
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor18.Appearance = appearance51;
			this.ultraNumericEditor18.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "l3"));
			this.ultraNumericEditor18.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor18.FormatString = "";
			this.ultraNumericEditor18.Location = new System.Drawing.Point(184, 432);
			this.ultraNumericEditor18.MaskInput = "nnnnnnn";
			this.ultraNumericEditor18.Name = "ultraNumericEditor18";
			this.ultraNumericEditor18.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor18.PromptChar = ' ';
			this.ultraNumericEditor18.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor18.TabIndex = 40;
			this.ultraNumericEditor18.Enter += new System.EventHandler(this.ultraNumericEditor18_Enter);
			// 
			// ultraNumericEditor19
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor19.Appearance = appearance52;
			this.ultraNumericEditor19.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "l2"));
			this.ultraNumericEditor19.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor19.FormatString = "";
			this.ultraNumericEditor19.Location = new System.Drawing.Point(184, 408);
			this.ultraNumericEditor19.MaskInput = "nnnnnnn";
			this.ultraNumericEditor19.Name = "ultraNumericEditor19";
			this.ultraNumericEditor19.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor19.PromptChar = ' ';
			this.ultraNumericEditor19.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor19.TabIndex = 37;
			this.ultraNumericEditor19.Enter += new System.EventHandler(this.ultraNumericEditor19_Enter);
			// 
			// ultraNumericEditor20
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor20.Appearance = appearance53;
			this.ultraNumericEditor20.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "l1"));
			this.ultraNumericEditor20.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor20.FormatString = "";
			this.ultraNumericEditor20.Location = new System.Drawing.Point(184, 384);
			this.ultraNumericEditor20.MaskInput = "nnnnnnn";
			this.ultraNumericEditor20.Name = "ultraNumericEditor20";
			this.ultraNumericEditor20.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor20.PromptChar = ' ';
			this.ultraNumericEditor20.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor20.TabIndex = 34;
			this.ultraNumericEditor20.Enter += new System.EventHandler(this.ultraNumericEditor20_Enter);
			// 
			// ultraNumericEditor22
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor22.Appearance = appearance54;
			this.ultraNumericEditor22.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "l8"));
			this.ultraNumericEditor22.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor22.FormatString = "";
			this.ultraNumericEditor22.Location = new System.Drawing.Point(184, 336);
			this.ultraNumericEditor22.MaskInput = "nnnnnnn";
			this.ultraNumericEditor22.Name = "ultraNumericEditor22";
			this.ultraNumericEditor22.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor22.PromptChar = ' ';
			this.ultraNumericEditor22.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor22.TabIndex = 28;
			this.ultraNumericEditor22.Enter += new System.EventHandler(this.ultraNumericEditor22_Enter);
			// 
			// ultraNumericEditor23
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor23.Appearance = appearance55;
			this.ultraNumericEditor23.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "v10"));
			this.ultraNumericEditor23.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor23.FormatString = "#,##0.00";
			this.ultraNumericEditor23.Location = new System.Drawing.Point(248, 552);
			this.ultraNumericEditor23.Name = "ultraNumericEditor23";
			this.ultraNumericEditor23.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor23.PromptChar = ' ';
			this.ultraNumericEditor23.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor23.TabIndex = 56;
			this.ultraNumericEditor23.Validated += new System.EventHandler(this.ultraNumericEditor23_Validated);
			this.ultraNumericEditor23.Enter += new System.EventHandler(this.ultraNumericEditor23_Enter);
			// 
			// ultraNumericEditor24
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor24.Appearance = appearance56;
			this.ultraNumericEditor24.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "v9"));
			this.ultraNumericEditor24.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor24.FormatString = "#,##0.00";
			this.ultraNumericEditor24.Location = new System.Drawing.Point(248, 528);
			this.ultraNumericEditor24.Name = "ultraNumericEditor24";
			this.ultraNumericEditor24.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor24.PromptChar = ' ';
			this.ultraNumericEditor24.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor24.TabIndex = 53;
			this.ultraNumericEditor24.Validated += new System.EventHandler(this.ultraNumericEditor24_Validated);
			this.ultraNumericEditor24.Enter += new System.EventHandler(this.ultraNumericEditor24_Enter);
			// 
			// ultraNumericEditor25
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor25.Appearance = appearance57;
			this.ultraNumericEditor25.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "v6"));
			this.ultraNumericEditor25.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor25.FormatString = "#,##0.00";
			this.ultraNumericEditor25.Location = new System.Drawing.Point(248, 504);
			this.ultraNumericEditor25.Name = "ultraNumericEditor25";
			this.ultraNumericEditor25.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor25.PromptChar = ' ';
			this.ultraNumericEditor25.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor25.TabIndex = 50;
			this.ultraNumericEditor25.Validated += new System.EventHandler(this.ultraNumericEditor25_Validated);
			this.ultraNumericEditor25.Enter += new System.EventHandler(this.ultraNumericEditor25_Enter);
			// 
			// ultraNumericEditor26
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor26.Appearance = appearance58;
			this.ultraNumericEditor26.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "v5"));
			this.ultraNumericEditor26.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor26.FormatString = "#,##0.00";
			this.ultraNumericEditor26.Location = new System.Drawing.Point(248, 480);
			this.ultraNumericEditor26.Name = "ultraNumericEditor26";
			this.ultraNumericEditor26.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor26.PromptChar = ' ';
			this.ultraNumericEditor26.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor26.TabIndex = 47;
			this.ultraNumericEditor26.Validated += new System.EventHandler(this.ultraNumericEditor26_Validated);
			this.ultraNumericEditor26.Enter += new System.EventHandler(this.ultraNumericEditor26_Enter);
			// 
			// ultraNumericEditor27
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor27.Appearance = appearance59;
			this.ultraNumericEditor27.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "v7"));
			this.ultraNumericEditor27.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor27.FormatString = "#,##0.00";
			this.ultraNumericEditor27.Location = new System.Drawing.Point(248, 360);
			this.ultraNumericEditor27.Name = "ultraNumericEditor27";
			this.ultraNumericEditor27.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor27.PromptChar = ' ';
			this.ultraNumericEditor27.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor27.TabIndex = 32;
			this.ultraNumericEditor27.Validated += new System.EventHandler(this.ultraNumericEditor27_Validated);
			this.ultraNumericEditor27.Enter += new System.EventHandler(this.ultraNumericEditor27_Enter);
			// 
			// ultraNumericEditor28
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor28.Appearance = appearance60;
			this.ultraNumericEditor28.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "v4"));
			this.ultraNumericEditor28.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor28.FormatString = "#,##0.00";
			this.ultraNumericEditor28.Location = new System.Drawing.Point(248, 456);
			this.ultraNumericEditor28.Name = "ultraNumericEditor28";
			this.ultraNumericEditor28.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor28.PromptChar = ' ';
			this.ultraNumericEditor28.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor28.TabIndex = 44;
			this.ultraNumericEditor28.Validated += new System.EventHandler(this.ultraNumericEditor28_Validated);
			this.ultraNumericEditor28.Enter += new System.EventHandler(this.ultraNumericEditor28_Enter);
			// 
			// ultraNumericEditor29
			// 
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor29.Appearance = appearance61;
			this.ultraNumericEditor29.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "v3"));
			this.ultraNumericEditor29.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor29.FormatString = "#,##0.00";
			this.ultraNumericEditor29.Location = new System.Drawing.Point(248, 432);
			this.ultraNumericEditor29.Name = "ultraNumericEditor29";
			this.ultraNumericEditor29.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor29.PromptChar = ' ';
			this.ultraNumericEditor29.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor29.TabIndex = 41;
			this.ultraNumericEditor29.Validated += new System.EventHandler(this.ultraNumericEditor29_Validated);
			this.ultraNumericEditor29.Enter += new System.EventHandler(this.ultraNumericEditor29_Enter);
			// 
			// ultraNumericEditor30
			// 
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor30.Appearance = appearance62;
			this.ultraNumericEditor30.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "v2"));
			this.ultraNumericEditor30.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor30.FormatString = "#,##0.00";
			this.ultraNumericEditor30.Location = new System.Drawing.Point(248, 408);
			this.ultraNumericEditor30.Name = "ultraNumericEditor30";
			this.ultraNumericEditor30.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor30.PromptChar = ' ';
			this.ultraNumericEditor30.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor30.TabIndex = 38;
			this.ultraNumericEditor30.Validated += new System.EventHandler(this.ultraNumericEditor30_Validated);
			this.ultraNumericEditor30.Enter += new System.EventHandler(this.ultraNumericEditor30_Enter);
			// 
			// ultraNumericEditor31
			// 
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor31.Appearance = appearance63;
			this.ultraNumericEditor31.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "v1"));
			this.ultraNumericEditor31.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor31.FormatString = "#,##0.00";
			this.ultraNumericEditor31.Location = new System.Drawing.Point(248, 384);
			this.ultraNumericEditor31.Name = "ultraNumericEditor31";
			this.ultraNumericEditor31.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor31.PromptChar = ' ';
			this.ultraNumericEditor31.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor31.TabIndex = 35;
			this.ultraNumericEditor31.Validated += new System.EventHandler(this.ultraNumericEditor31_Validated);
			this.ultraNumericEditor31.Enter += new System.EventHandler(this.ultraNumericEditor31_Enter);
			// 
			// ultraNumericEditor33
			// 
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor33.Appearance = appearance64;
			this.ultraNumericEditor33.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "v8"));
			this.ultraNumericEditor33.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor33.FormatString = "#,##0.00";
			this.ultraNumericEditor33.Location = new System.Drawing.Point(248, 336);
			this.ultraNumericEditor33.Name = "ultraNumericEditor33";
			this.ultraNumericEditor33.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor33.PromptChar = ' ';
			this.ultraNumericEditor33.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor33.TabIndex = 29;
			this.ultraNumericEditor33.Validated += new System.EventHandler(this.ultraNumericEditor33_Validated);
			this.ultraNumericEditor33.ValueChanged += new System.EventHandler(this.ultraNumericEditor33_ValueChanged);
			this.ultraNumericEditor33.Enter += new System.EventHandler(this.ultraNumericEditor33_Enter);
			// 
			// label83
			// 
			this.label83.AutoSize = true;
			this.label83.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label83.Location = new System.Drawing.Point(136, 304);
			this.label83.Name = "label83";
			this.label83.Size = new System.Drawing.Size(156, 18);
			this.label83.TabIndex = 206;
			this.label83.Text = "E L E C T R O N I C O S";
			// 
			// ultraNumericEditor34
			// 
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor34.Appearance = appearance65;
			this.ultraNumericEditor34.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb10"));
			this.ultraNumericEditor34.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor34.FormatString = "";
			this.ultraNumericEditor34.Location = new System.Drawing.Point(512, 288);
			this.ultraNumericEditor34.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor34.Name = "ultraNumericEditor34";
			this.ultraNumericEditor34.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor34.PromptChar = ' ';
			this.ultraNumericEditor34.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor34.TabIndex = 85;
			this.ultraNumericEditor34.Enter += new System.EventHandler(this.ultraNumericEditor34_Enter);
			// 
			// ultraNumericEditor38
			// 
			appearance66.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor38.Appearance = appearance66;
			this.ultraNumericEditor38.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb6"));
			this.ultraNumericEditor38.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor38.FormatString = "";
			this.ultraNumericEditor38.Location = new System.Drawing.Point(512, 192);
			this.ultraNumericEditor38.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor38.Name = "ultraNumericEditor38";
			this.ultraNumericEditor38.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor38.PromptChar = ' ';
			this.ultraNumericEditor38.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor38.TabIndex = 73;
			this.ultraNumericEditor38.ValueChanged += new System.EventHandler(this.ultraNumericEditor38_ValueChanged);
			this.ultraNumericEditor38.Enter += new System.EventHandler(this.ultraNumericEditor38_Enter);
			// 
			// ultraNumericEditor39
			// 
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor39.Appearance = appearance67;
			this.ultraNumericEditor39.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb5"));
			this.ultraNumericEditor39.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor39.FormatString = "";
			this.ultraNumericEditor39.Location = new System.Drawing.Point(512, 168);
			this.ultraNumericEditor39.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor39.Name = "ultraNumericEditor39";
			this.ultraNumericEditor39.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor39.PromptChar = ' ';
			this.ultraNumericEditor39.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor39.TabIndex = 70;
			this.ultraNumericEditor39.ValueChanged += new System.EventHandler(this.ultraNumericEditor39_ValueChanged);
			this.ultraNumericEditor39.Enter += new System.EventHandler(this.ultraNumericEditor39_Enter);
			// 
			// ultraNumericEditor40
			// 
			appearance68.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor40.Appearance = appearance68;
			this.ultraNumericEditor40.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb4"));
			this.ultraNumericEditor40.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor40.FormatString = "";
			this.ultraNumericEditor40.Location = new System.Drawing.Point(512, 144);
			this.ultraNumericEditor40.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor40.Name = "ultraNumericEditor40";
			this.ultraNumericEditor40.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor40.PromptChar = ' ';
			this.ultraNumericEditor40.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor40.TabIndex = 67;
			this.ultraNumericEditor40.Enter += new System.EventHandler(this.ultraNumericEditor40_Enter);
			// 
			// ultraNumericEditor41
			// 
			appearance69.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor41.Appearance = appearance69;
			this.ultraNumericEditor41.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb3"));
			this.ultraNumericEditor41.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor41.FormatString = "";
			this.ultraNumericEditor41.Location = new System.Drawing.Point(512, 120);
			this.ultraNumericEditor41.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor41.Name = "ultraNumericEditor41";
			this.ultraNumericEditor41.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor41.PromptChar = ' ';
			this.ultraNumericEditor41.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor41.TabIndex = 64;
			this.ultraNumericEditor41.ValueChanged += new System.EventHandler(this.ultraNumericEditor41_ValueChanged);
			this.ultraNumericEditor41.Enter += new System.EventHandler(this.ultraNumericEditor41_Enter);
			// 
			// ultraNumericEditor42
			// 
			appearance70.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor42.Appearance = appearance70;
			this.ultraNumericEditor42.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb2"));
			this.ultraNumericEditor42.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor42.FormatString = "";
			this.ultraNumericEditor42.Location = new System.Drawing.Point(512, 96);
			this.ultraNumericEditor42.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor42.Name = "ultraNumericEditor42";
			this.ultraNumericEditor42.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor42.PromptChar = ' ';
			this.ultraNumericEditor42.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor42.TabIndex = 61;
			this.ultraNumericEditor42.ValueChanged += new System.EventHandler(this.ultraNumericEditor42_ValueChanged);
			this.ultraNumericEditor42.Enter += new System.EventHandler(this.ultraNumericEditor42_Enter);
			// 
			// ultraNumericEditor45
			// 
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor45.Appearance = appearance71;
			this.ultraNumericEditor45.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv10"));
			this.ultraNumericEditor45.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor45.FormatString = "#,##0.00";
			this.ultraNumericEditor45.Location = new System.Drawing.Point(592, 288);
			this.ultraNumericEditor45.Name = "ultraNumericEditor45";
			this.ultraNumericEditor45.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor45.PromptChar = ' ';
			this.ultraNumericEditor45.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor45.TabIndex = 86;
			this.ultraNumericEditor45.Validated += new System.EventHandler(this.ultraNumericEditor45_Validated);
			this.ultraNumericEditor45.ValueChanged += new System.EventHandler(this.ultraNumericEditor45_ValueChanged);
			this.ultraNumericEditor45.Enter += new System.EventHandler(this.ultraNumericEditor45_Enter);
			// 
			// ultraNumericEditor49
			// 
			appearance72.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor49.Appearance = appearance72;
			this.ultraNumericEditor49.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv6"));
			this.ultraNumericEditor49.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor49.FormatString = "#,##0.00";
			this.ultraNumericEditor49.Location = new System.Drawing.Point(592, 192);
			this.ultraNumericEditor49.Name = "ultraNumericEditor49";
			this.ultraNumericEditor49.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor49.PromptChar = ' ';
			this.ultraNumericEditor49.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor49.TabIndex = 74;
			this.ultraNumericEditor49.Validated += new System.EventHandler(this.ultraNumericEditor49_Validated);
			this.ultraNumericEditor49.Enter += new System.EventHandler(this.ultraNumericEditor49_Enter);
			// 
			// ultraNumericEditor50
			// 
			appearance73.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor50.Appearance = appearance73;
			this.ultraNumericEditor50.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv5"));
			this.ultraNumericEditor50.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor50.FormatString = "#,##0.00";
			this.ultraNumericEditor50.Location = new System.Drawing.Point(592, 168);
			this.ultraNumericEditor50.Name = "ultraNumericEditor50";
			this.ultraNumericEditor50.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor50.PromptChar = ' ';
			this.ultraNumericEditor50.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor50.TabIndex = 71;
			this.ultraNumericEditor50.Validated += new System.EventHandler(this.ultraNumericEditor50_Validated);
			this.ultraNumericEditor50.Enter += new System.EventHandler(this.ultraNumericEditor50_Enter);
			// 
			// ultraNumericEditor51
			// 
			appearance74.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor51.Appearance = appearance74;
			this.ultraNumericEditor51.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv4"));
			this.ultraNumericEditor51.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor51.FormatString = "#,##0.00";
			this.ultraNumericEditor51.Location = new System.Drawing.Point(592, 144);
			this.ultraNumericEditor51.Name = "ultraNumericEditor51";
			this.ultraNumericEditor51.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor51.PromptChar = ' ';
			this.ultraNumericEditor51.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor51.TabIndex = 68;
			this.ultraNumericEditor51.Validated += new System.EventHandler(this.ultraNumericEditor51_Validated);
			this.ultraNumericEditor51.ValueChanged += new System.EventHandler(this.ultraNumericEditor51_ValueChanged);
			this.ultraNumericEditor51.Enter += new System.EventHandler(this.ultraNumericEditor51_Enter);
			// 
			// ultraNumericEditor52
			// 
			appearance75.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor52.Appearance = appearance75;
			this.ultraNumericEditor52.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv3"));
			this.ultraNumericEditor52.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor52.FormatString = "#,##0.00";
			this.ultraNumericEditor52.Location = new System.Drawing.Point(592, 120);
			this.ultraNumericEditor52.Name = "ultraNumericEditor52";
			this.ultraNumericEditor52.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor52.PromptChar = ' ';
			this.ultraNumericEditor52.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor52.TabIndex = 65;
			this.ultraNumericEditor52.Validated += new System.EventHandler(this.ultraNumericEditor52_Validated);
			this.ultraNumericEditor52.Enter += new System.EventHandler(this.ultraNumericEditor52_Enter);
			// 
			// ultraNumericEditor53
			// 
			appearance76.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor53.Appearance = appearance76;
			this.ultraNumericEditor53.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv2"));
			this.ultraNumericEditor53.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor53.FormatString = "#,##0.00";
			this.ultraNumericEditor53.Location = new System.Drawing.Point(592, 96);
			this.ultraNumericEditor53.Name = "ultraNumericEditor53";
			this.ultraNumericEditor53.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor53.PromptChar = ' ';
			this.ultraNumericEditor53.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor53.TabIndex = 62;
			this.ultraNumericEditor53.Validated += new System.EventHandler(this.ultraNumericEditor53_Validated);
			this.ultraNumericEditor53.Enter += new System.EventHandler(this.ultraNumericEditor53_Enter);
			// 
			// label86
			// 
			this.label86.AutoSize = true;
			this.label86.Location = new System.Drawing.Point(136, 320);
			this.label86.Name = "label86";
			this.label86.Size = new System.Drawing.Size(35, 16);
			this.label86.TabIndex = 242;
			this.label86.Text = "CANT";
			// 
			// label87
			// 
			this.label87.AutoSize = true;
			this.label87.Location = new System.Drawing.Point(192, 320);
			this.label87.Name = "label87";
			this.label87.Size = new System.Drawing.Size(34, 16);
			this.label87.TabIndex = 243;
			this.label87.Text = "LOTE";
			// 
			// label88
			// 
			this.label88.AutoSize = true;
			this.label88.Location = new System.Drawing.Point(256, 320);
			this.label88.Name = "label88";
			this.label88.Size = new System.Drawing.Size(43, 16);
			this.label88.TabIndex = 244;
			this.label88.Text = "VALOR";
			this.label88.Click += new System.EventHandler(this.label88_Click);
			// 
			// txtIniciaF
			// 
			appearance77.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIniciaF.Appearance = appearance77;
			this.txtIniciaF.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Inicia"));
			this.txtIniciaF.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIniciaF.FormatString = "";
			this.txtIniciaF.Location = new System.Drawing.Point(760, 504);
			this.txtIniciaF.MaskInput = "nnnnnnnnn";
			this.txtIniciaF.Name = "txtIniciaF";
			this.txtIniciaF.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIniciaF.PromptChar = ' ';
			this.txtIniciaF.Size = new System.Drawing.Size(88, 21);
			this.txtIniciaF.TabIndex = 144;
			this.txtIniciaF.Enter += new System.EventHandler(this.txtIniciaF_Enter);
			// 
			// txtTerminaF
			// 
			appearance78.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTerminaF.Appearance = appearance78;
			this.txtTerminaF.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Termina"));
			this.txtTerminaF.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTerminaF.FormatString = "";
			this.txtTerminaF.Location = new System.Drawing.Point(760, 528);
			this.txtTerminaF.MaskInput = "nnnnnnnnn";
			this.txtTerminaF.Name = "txtTerminaF";
			this.txtTerminaF.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTerminaF.PromptChar = ' ';
			this.txtTerminaF.Size = new System.Drawing.Size(88, 21);
			this.txtTerminaF.TabIndex = 145;
			this.txtTerminaF.Enter += new System.EventHandler(this.txtTerminaF_Enter);
			// 
			// label91
			// 
			this.label91.AutoSize = true;
			this.label91.Location = new System.Drawing.Point(720, 504);
			this.label91.Name = "label91";
			this.label91.Size = new System.Drawing.Size(31, 16);
			this.label91.TabIndex = 260;
			this.label91.Text = "Inicia";
			// 
			// label92
			// 
			this.label92.AutoSize = true;
			this.label92.Location = new System.Drawing.Point(712, 536);
			this.label92.Name = "label92";
			this.label92.Size = new System.Drawing.Size(46, 16);
			this.label92.TabIndex = 261;
			this.label92.Text = "Termina";
			// 
			// label93
			// 
			this.label93.AutoSize = true;
			this.label93.Location = new System.Drawing.Point(912, 264);
			this.label93.Name = "label93";
			this.label93.Size = new System.Drawing.Size(53, 16);
			this.label93.TabIndex = 262;
			this.label93.Text = "Promo 35";
			// 
			// label94
			// 
			this.label94.AutoSize = true;
			this.label94.Location = new System.Drawing.Point(912, 288);
			this.label94.Name = "label94";
			this.label94.Size = new System.Drawing.Size(53, 16);
			this.label94.TabIndex = 263;
			this.label94.Text = "Promo 79";
			// 
			// label95
			// 
			this.label95.AutoSize = true;
			this.label95.Location = new System.Drawing.Point(200, 120);
			this.label95.Name = "label95";
			this.label95.Size = new System.Drawing.Size(55, 16);
			this.label95.TabIndex = 264;
			this.label95.Text = "Retención";
			// 
			// label96
			// 
			this.label96.AutoSize = true;
			this.label96.Location = new System.Drawing.Point(16, 144);
			this.label96.Name = "label96";
			this.label96.Size = new System.Drawing.Size(77, 16);
			this.label96.TabIndex = 265;
			this.label96.Text = "Cred Empresa";
			this.label96.Click += new System.EventHandler(this.label96_Click);
			// 
			// txtPromo35
			// 
			appearance79.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPromo35.Appearance = appearance79;
			this.txtPromo35.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Promocion35"));
			this.txtPromo35.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPromo35.FormatString = "0";
			this.txtPromo35.Location = new System.Drawing.Point(968, 256);
			this.txtPromo35.MaskInput = "nnn";
			this.txtPromo35.Name = "txtPromo35";
			this.txtPromo35.PromptChar = ' ';
			this.txtPromo35.Size = new System.Drawing.Size(32, 21);
			this.txtPromo35.TabIndex = 140;
			this.txtPromo35.Validated += new System.EventHandler(this.txtPromo35_Validated);
			this.txtPromo35.Enter += new System.EventHandler(this.txtPromo35_Enter);
			// 
			// txtPromo79
			// 
			appearance80.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPromo79.Appearance = appearance80;
			this.txtPromo79.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Promocion79"));
			this.txtPromo79.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPromo79.FormatString = "0";
			this.txtPromo79.Location = new System.Drawing.Point(968, 280);
			this.txtPromo79.MaskInput = "nnn";
			this.txtPromo79.Name = "txtPromo79";
			this.txtPromo79.PromptChar = ' ';
			this.txtPromo79.Size = new System.Drawing.Size(32, 21);
			this.txtPromo79.TabIndex = 141;
			this.txtPromo79.Validated += new System.EventHandler(this.txtPromo79_Validated);
			this.txtPromo79.Enter += new System.EventHandler(this.txtPromo79_Enter);
			// 
			// txtRetencion
			// 
			appearance81.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetencion.Appearance = appearance81;
			this.txtRetencion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "Retencion"));
			this.txtRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetencion.FormatString = "#,##0.00";
			this.txtRetencion.Location = new System.Drawing.Point(272, 112);
			this.txtRetencion.Name = "txtRetencion";
			this.txtRetencion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetencion.PromptChar = ' ';
			this.txtRetencion.Size = new System.Drawing.Size(96, 21);
			this.txtRetencion.TabIndex = 9;
			this.txtRetencion.Validated += new System.EventHandler(this.txtRetencion_Validated);
			this.txtRetencion.Enter += new System.EventHandler(this.txtRetencion_Enter);
			// 
			// txtCredEmpresa
			// 
			appearance82.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCredEmpresa.Appearance = appearance82;
			this.txtCredEmpresa.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "CreditoEmpresa"));
			this.txtCredEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCredEmpresa.FormatString = "#,##0.00";
			this.txtCredEmpresa.Location = new System.Drawing.Point(104, 136);
			this.txtCredEmpresa.Name = "txtCredEmpresa";
			this.txtCredEmpresa.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCredEmpresa.PromptChar = ' ';
			this.txtCredEmpresa.Size = new System.Drawing.Size(80, 21);
			this.txtCredEmpresa.TabIndex = 10;
			this.txtCredEmpresa.Validated += new System.EventHandler(this.txtCredEmpresa_Validated);
			this.txtCredEmpresa.Enter += new System.EventHandler(this.txtCredEmpresa_Enter);
			// 
			// label97
			// 
			this.label97.AutoSize = true;
			this.label97.Location = new System.Drawing.Point(856, 536);
			this.label97.Name = "label97";
			this.label97.Size = new System.Drawing.Size(65, 16);
			this.label97.TabIndex = 271;
			this.label97.Text = "Termina NC";
			// 
			// label98
			// 
			this.label98.AutoSize = true;
			this.label98.Location = new System.Drawing.Point(856, 512);
			this.label98.Name = "label98";
			this.label98.Size = new System.Drawing.Size(50, 16);
			this.label98.TabIndex = 270;
			this.label98.Text = "Inicia NC";
			// 
			// txtTerminaNC
			// 
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTerminaNC.Appearance = appearance83;
			this.txtTerminaNC.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "TerminaNC"));
			this.txtTerminaNC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTerminaNC.FormatString = "";
			this.txtTerminaNC.Location = new System.Drawing.Point(928, 528);
			this.txtTerminaNC.MaskInput = "nnnnnnnnn";
			this.txtTerminaNC.Name = "txtTerminaNC";
			this.txtTerminaNC.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTerminaNC.PromptChar = ' ';
			this.txtTerminaNC.Size = new System.Drawing.Size(72, 21);
			this.txtTerminaNC.TabIndex = 147;
			this.txtTerminaNC.Enter += new System.EventHandler(this.txtTerminaNC_Enter);
			// 
			// txtIniciaNC
			// 
			appearance84.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIniciaNC.Appearance = appearance84;
			this.txtIniciaNC.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "IniciaNC"));
			this.txtIniciaNC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIniciaNC.FormatString = "";
			this.txtIniciaNC.Location = new System.Drawing.Point(928, 504);
			this.txtIniciaNC.MaskInput = "nnnnnnnnn";
			this.txtIniciaNC.Name = "txtIniciaNC";
			this.txtIniciaNC.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIniciaNC.PromptChar = ' ';
			this.txtIniciaNC.Size = new System.Drawing.Size(72, 21);
			this.txtIniciaNC.TabIndex = 146;
			this.txtIniciaNC.Enter += new System.EventHandler(this.txtIniciaNC_Enter);
			// 
			// ultraNumericEditor48
			// 
			appearance85.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor48.Appearance = appearance85;
			this.ultraNumericEditor48.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv7"));
			this.ultraNumericEditor48.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor48.FormatString = "#,##0.00";
			this.ultraNumericEditor48.Location = new System.Drawing.Point(592, 216);
			this.ultraNumericEditor48.Name = "ultraNumericEditor48";
			this.ultraNumericEditor48.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor48.PromptChar = ' ';
			this.ultraNumericEditor48.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor48.TabIndex = 77;
			this.ultraNumericEditor48.Validated += new System.EventHandler(this.ultraNumericEditor48_Validated);
			this.ultraNumericEditor48.ValueChanged += new System.EventHandler(this.ultraNumericEditor48_ValueChanged);
			this.ultraNumericEditor48.Enter += new System.EventHandler(this.ultraNumericEditor48_Enter);
			// 
			// ultraNumericEditor37
			// 
			appearance86.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor37.Appearance = appearance86;
			this.ultraNumericEditor37.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb7"));
			this.ultraNumericEditor37.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor37.FormatString = "";
			this.ultraNumericEditor37.Location = new System.Drawing.Point(512, 216);
			this.ultraNumericEditor37.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor37.Name = "ultraNumericEditor37";
			this.ultraNumericEditor37.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor37.PromptChar = ' ';
			this.ultraNumericEditor37.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor37.TabIndex = 76;
			this.ultraNumericEditor37.Enter += new System.EventHandler(this.ultraNumericEditor37_Enter);
			// 
			// ultraNumericEditor36
			// 
			appearance87.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor36.Appearance = appearance87;
			this.ultraNumericEditor36.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb8"));
			this.ultraNumericEditor36.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor36.FormatString = "";
			this.ultraNumericEditor36.Location = new System.Drawing.Point(512, 240);
			this.ultraNumericEditor36.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor36.Name = "ultraNumericEditor36";
			this.ultraNumericEditor36.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor36.PromptChar = ' ';
			this.ultraNumericEditor36.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor36.TabIndex = 79;
			this.ultraNumericEditor36.Enter += new System.EventHandler(this.ultraNumericEditor36_Enter);
			// 
			// ultraNumericEditor47
			// 
			appearance88.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor47.Appearance = appearance88;
			this.ultraNumericEditor47.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv8"));
			this.ultraNumericEditor47.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor47.FormatString = "#,##0.00";
			this.ultraNumericEditor47.Location = new System.Drawing.Point(592, 240);
			this.ultraNumericEditor47.Name = "ultraNumericEditor47";
			this.ultraNumericEditor47.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor47.PromptChar = ' ';
			this.ultraNumericEditor47.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor47.TabIndex = 80;
			this.ultraNumericEditor47.Validated += new System.EventHandler(this.ultraNumericEditor47_Validated);
			this.ultraNumericEditor47.Enter += new System.EventHandler(this.ultraNumericEditor47_Enter);
			// 
			// ultraNumericEditor46
			// 
			appearance89.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor46.Appearance = appearance89;
			this.ultraNumericEditor46.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv9"));
			this.ultraNumericEditor46.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor46.FormatString = "#,##0.00";
			this.ultraNumericEditor46.Location = new System.Drawing.Point(592, 264);
			this.ultraNumericEditor46.Name = "ultraNumericEditor46";
			this.ultraNumericEditor46.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor46.PromptChar = ' ';
			this.ultraNumericEditor46.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor46.TabIndex = 83;
			this.ultraNumericEditor46.Validated += new System.EventHandler(this.ultraNumericEditor46_Validated);
			this.ultraNumericEditor46.ValueChanged += new System.EventHandler(this.ultraNumericEditor46_ValueChanged);
			this.ultraNumericEditor46.Enter += new System.EventHandler(this.ultraNumericEditor46_Enter);
			// 
			// ultraNumericEditor35
			// 
			appearance90.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor35.Appearance = appearance90;
			this.ultraNumericEditor35.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb9"));
			this.ultraNumericEditor35.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor35.FormatString = "";
			this.ultraNumericEditor35.Location = new System.Drawing.Point(512, 264);
			this.ultraNumericEditor35.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor35.Name = "ultraNumericEditor35";
			this.ultraNumericEditor35.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor35.PromptChar = ' ';
			this.ultraNumericEditor35.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor35.TabIndex = 82;
			this.ultraNumericEditor35.ValueChanged += new System.EventHandler(this.ultraNumericEditor35_ValueChanged);
			this.ultraNumericEditor35.Enter += new System.EventHandler(this.ultraNumericEditor35_Enter);
			// 
			// ultraNumericEditor2
			// 
			appearance91.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor2.Appearance = appearance91;
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb1"));
			this.ultraNumericEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor2.FormatString = "";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(512, 72);
			this.ultraNumericEditor2.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor2.TabIndex = 58;
			// 
			// ultraNumericEditor21
			// 
			appearance92.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor21.Appearance = appearance92;
			this.ultraNumericEditor21.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv1"));
			this.ultraNumericEditor21.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor21.FormatString = "#,##0.00";
			this.ultraNumericEditor21.Location = new System.Drawing.Point(592, 72);
			this.ultraNumericEditor21.Name = "ultraNumericEditor21";
			this.ultraNumericEditor21.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor21.PromptChar = ' ';
			this.ultraNumericEditor21.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor21.TabIndex = 59;
			this.ultraNumericEditor21.Validated += new System.EventHandler(this.ultraNumericEditor21_Validated);
			this.ultraNumericEditor21.Enter += new System.EventHandler(this.ultraNumericEditor21_Enter_1);
			// 
			// cdsFormaPago
			// 
			this.cdsFormaPago.BindingContextCtrl = this;
			this.cdsFormaPago.DataLibrary = "LibFacturacion";
			this.cdsFormaPago.DataLibraryUrl = "";
			this.cdsFormaPago.DataSetDef = "dsCompraFormaPago2";
			this.cdsFormaPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsFormaPago.Name = "cdsFormaPago";
			this.cdsFormaPago.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsFormaPago.SchemaDef = null;
			this.cdsFormaPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsFormaPago_BeforeFill);
			// 
			// cdvFormaPago
			// 
			this.cdvFormaPago.BindingContextCtrl = this;
			this.cdvFormaPago.DataSet = this.cdsFormaPago;
			this.cdvFormaPago.TableName = "";
			this.cdvFormaPago.TableViewName = "CompraFormaPago2";
			// 
			// ultraCombo1
			// 
			appearance93.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo1.Appearance = appearance93;
			this.ultraCombo1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt1"));
			this.ultraCombo1.DataSource = this.cdvFormaPago;
			ultraGridBand2.AddButtonCaption = "MarceloBodega";
			this.ultraCombo1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraCombo1.DisplayMember = "FormaPago";
			this.ultraCombo1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo1.Location = new System.Drawing.Point(384, 72);
			this.ultraCombo1.Name = "ultraCombo1";
			this.ultraCombo1.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo1.TabIndex = 57;
			this.ultraCombo1.ValueMember = "idFormaPago";
			this.ultraCombo1.Validated += new System.EventHandler(this.ultraCombo1_Validated);
			// 
			// ultraCombo2
			// 
			appearance94.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo2.Appearance = appearance94;
			this.ultraCombo2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt2"));
			this.ultraCombo2.DataSource = this.cdvFormaPago;
			ultraGridBand3.AddButtonCaption = "MarceloBodega";
			this.ultraCombo2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraCombo2.DisplayMember = "FormaPago";
			this.ultraCombo2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo2.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo2.Location = new System.Drawing.Point(384, 96);
			this.ultraCombo2.Name = "ultraCombo2";
			this.ultraCombo2.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo2.TabIndex = 60;
			this.ultraCombo2.ValueMember = "idFormaPago";
			this.ultraCombo2.Validated += new System.EventHandler(this.ultraCombo2_Validated);
			// 
			// ultraCombo3
			// 
			appearance95.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo3.Appearance = appearance95;
			this.ultraCombo3.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo3.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt3"));
			this.ultraCombo3.DataSource = this.cdvFormaPago;
			ultraGridBand4.AddButtonCaption = "MarceloBodega";
			this.ultraCombo3.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraCombo3.DisplayMember = "FormaPago";
			this.ultraCombo3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo3.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo3.Location = new System.Drawing.Point(384, 120);
			this.ultraCombo3.Name = "ultraCombo3";
			this.ultraCombo3.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo3.TabIndex = 63;
			this.ultraCombo3.ValueMember = "idFormaPago";
			this.ultraCombo3.Validated += new System.EventHandler(this.ultraCombo3_Validated);
			// 
			// ultraCombo4
			// 
			appearance96.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo4.Appearance = appearance96;
			this.ultraCombo4.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo4.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt4"));
			this.ultraCombo4.DataSource = this.cdvFormaPago;
			ultraGridBand5.AddButtonCaption = "MarceloBodega";
			this.ultraCombo4.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.ultraCombo4.DisplayMember = "FormaPago";
			this.ultraCombo4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo4.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo4.Location = new System.Drawing.Point(384, 144);
			this.ultraCombo4.Name = "ultraCombo4";
			this.ultraCombo4.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo4.TabIndex = 66;
			this.ultraCombo4.ValueMember = "idFormaPago";
			this.ultraCombo4.Validated += new System.EventHandler(this.ultraCombo4_Validated);
			// 
			// ultraCombo5
			// 
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo5.Appearance = appearance97;
			this.ultraCombo5.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo5.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt5"));
			this.ultraCombo5.DataSource = this.cdvFormaPago;
			ultraGridBand6.AddButtonCaption = "MarceloBodega";
			this.ultraCombo5.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ultraCombo5.DisplayMember = "FormaPago";
			this.ultraCombo5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo5.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo5.Location = new System.Drawing.Point(384, 168);
			this.ultraCombo5.Name = "ultraCombo5";
			this.ultraCombo5.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo5.TabIndex = 69;
			this.ultraCombo5.ValueMember = "idFormaPago";
			this.ultraCombo5.Validated += new System.EventHandler(this.ultraCombo5_Validated);
			// 
			// ultraCombo6
			// 
			appearance98.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo6.Appearance = appearance98;
			this.ultraCombo6.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo6.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt6"));
			this.ultraCombo6.DataSource = this.cdvFormaPago;
			ultraGridBand7.AddButtonCaption = "MarceloBodega";
			this.ultraCombo6.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.ultraCombo6.DisplayMember = "FormaPago";
			this.ultraCombo6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo6.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo6.Location = new System.Drawing.Point(384, 192);
			this.ultraCombo6.Name = "ultraCombo6";
			this.ultraCombo6.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo6.TabIndex = 72;
			this.ultraCombo6.ValueMember = "idFormaPago";
			this.ultraCombo6.Validated += new System.EventHandler(this.ultraCombo6_Validated);
			// 
			// ultraCombo7
			// 
			appearance99.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo7.Appearance = appearance99;
			this.ultraCombo7.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo7.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt7"));
			this.ultraCombo7.DataSource = this.cdvFormaPago;
			ultraGridBand8.AddButtonCaption = "MarceloBodega";
			this.ultraCombo7.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.ultraCombo7.DisplayMember = "FormaPago";
			this.ultraCombo7.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo7.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo7.Location = new System.Drawing.Point(384, 216);
			this.ultraCombo7.Name = "ultraCombo7";
			this.ultraCombo7.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo7.TabIndex = 75;
			this.ultraCombo7.ValueMember = "idFormaPago";
			this.ultraCombo7.Validated += new System.EventHandler(this.ultraCombo7_Validated);
			this.ultraCombo7.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo7_InitializeLayout);
			// 
			// ultraCombo8
			// 
			appearance100.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo8.Appearance = appearance100;
			this.ultraCombo8.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo8.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo8.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt8"));
			this.ultraCombo8.DataSource = this.cdvFormaPago;
			ultraGridBand9.AddButtonCaption = "MarceloBodega";
			this.ultraCombo8.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.ultraCombo8.DisplayMember = "FormaPago";
			this.ultraCombo8.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo8.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo8.Location = new System.Drawing.Point(384, 240);
			this.ultraCombo8.Name = "ultraCombo8";
			this.ultraCombo8.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo8.TabIndex = 78;
			this.ultraCombo8.ValueMember = "idFormaPago";
			this.ultraCombo8.Validated += new System.EventHandler(this.ultraCombo8_Validated);
			// 
			// ultraCombo9
			// 
			appearance101.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo9.Appearance = appearance101;
			this.ultraCombo9.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo9.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo9.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt9"));
			this.ultraCombo9.DataSource = this.cdvFormaPago;
			ultraGridBand10.AddButtonCaption = "MarceloBodega";
			this.ultraCombo9.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.ultraCombo9.DisplayMember = "FormaPago";
			this.ultraCombo9.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo9.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo9.Location = new System.Drawing.Point(384, 264);
			this.ultraCombo9.Name = "ultraCombo9";
			this.ultraCombo9.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo9.TabIndex = 81;
			this.ultraCombo9.ValueMember = "idFormaPago";
			this.ultraCombo9.Validated += new System.EventHandler(this.ultraCombo9_Validated);
			// 
			// ultraCombo10
			// 
			appearance102.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo10.Appearance = appearance102;
			this.ultraCombo10.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo10.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo10.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt10"));
			this.ultraCombo10.DataSource = this.cdvFormaPago;
			ultraGridBand11.AddButtonCaption = "MarceloBodega";
			this.ultraCombo10.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.ultraCombo10.DisplayMember = "FormaPago";
			this.ultraCombo10.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo10.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo10.Location = new System.Drawing.Point(384, 288);
			this.ultraCombo10.Name = "ultraCombo10";
			this.ultraCombo10.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo10.TabIndex = 84;
			this.ultraCombo10.ValueMember = "idFormaPago";
			this.ultraCombo10.Validated += new System.EventHandler(this.ultraCombo10_Validated);
			this.ultraCombo10.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo10_InitializeLayout);
			// 
			// ultraCombo11
			// 
			appearance103.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo11.Appearance = appearance103;
			this.ultraCombo11.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo11.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo11.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt11"));
			this.ultraCombo11.DataSource = this.cdvFormaPago;
			ultraGridBand12.AddButtonCaption = "MarceloBodega";
			this.ultraCombo11.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.ultraCombo11.DisplayMember = "FormaPago";
			this.ultraCombo11.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo11.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo11.Location = new System.Drawing.Point(384, 312);
			this.ultraCombo11.Name = "ultraCombo11";
			this.ultraCombo11.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo11.TabIndex = 87;
			this.ultraCombo11.ValueMember = "idFormaPago";
			this.ultraCombo11.Validated += new System.EventHandler(this.ultraCombo11_Validated);
			// 
			// ultraNumericEditor32
			// 
			appearance104.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor32.Appearance = appearance104;
			this.ultraNumericEditor32.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb11"));
			this.ultraNumericEditor32.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor32.FormatString = "";
			this.ultraNumericEditor32.Location = new System.Drawing.Point(512, 312);
			this.ultraNumericEditor32.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor32.Name = "ultraNumericEditor32";
			this.ultraNumericEditor32.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor32.PromptChar = ' ';
			this.ultraNumericEditor32.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor32.TabIndex = 88;
			// 
			// ultraNumericEditor43
			// 
			appearance105.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor43.Appearance = appearance105;
			this.ultraNumericEditor43.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv11"));
			this.ultraNumericEditor43.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor43.FormatString = "#,##0.00";
			this.ultraNumericEditor43.Location = new System.Drawing.Point(592, 312);
			this.ultraNumericEditor43.Name = "ultraNumericEditor43";
			this.ultraNumericEditor43.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor43.PromptChar = ' ';
			this.ultraNumericEditor43.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor43.TabIndex = 89;
			this.ultraNumericEditor43.Validated += new System.EventHandler(this.ultraNumericEditor43_Validated);
			this.ultraNumericEditor43.Enter += new System.EventHandler(this.ultraNumericEditor43_Enter);
			// 
			// ultraNumericEditor44
			// 
			appearance106.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor44.Appearance = appearance106;
			this.ultraNumericEditor44.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv12"));
			this.ultraNumericEditor44.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor44.FormatString = "#,##0.00";
			this.ultraNumericEditor44.Location = new System.Drawing.Point(592, 336);
			this.ultraNumericEditor44.Name = "ultraNumericEditor44";
			this.ultraNumericEditor44.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor44.PromptChar = ' ';
			this.ultraNumericEditor44.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor44.TabIndex = 92;
			this.ultraNumericEditor44.Validated += new System.EventHandler(this.ultraNumericEditor44_Validated);
			this.ultraNumericEditor44.Enter += new System.EventHandler(this.ultraNumericEditor44_Enter);
			// 
			// ultraNumericEditor54
			// 
			appearance107.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor54.Appearance = appearance107;
			this.ultraNumericEditor54.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb12"));
			this.ultraNumericEditor54.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor54.FormatString = "";
			this.ultraNumericEditor54.Location = new System.Drawing.Point(512, 336);
			this.ultraNumericEditor54.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor54.Name = "ultraNumericEditor54";
			this.ultraNumericEditor54.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor54.PromptChar = ' ';
			this.ultraNumericEditor54.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor54.TabIndex = 91;
			// 
			// ultraCombo12
			// 
			appearance108.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo12.Appearance = appearance108;
			this.ultraCombo12.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo12.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo12.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt12"));
			this.ultraCombo12.DataSource = this.cdvFormaPago;
			ultraGridBand13.AddButtonCaption = "MarceloBodega";
			this.ultraCombo12.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.ultraCombo12.DisplayMember = "FormaPago";
			this.ultraCombo12.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo12.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo12.Location = new System.Drawing.Point(384, 336);
			this.ultraCombo12.Name = "ultraCombo12";
			this.ultraCombo12.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo12.TabIndex = 90;
			this.ultraCombo12.ValueMember = "idFormaPago";
			this.ultraCombo12.Validated += new System.EventHandler(this.ultraCombo12_Validated);
			this.ultraCombo12.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo12_InitializeLayout);
			// 
			// ultraNumericEditor55
			// 
			appearance109.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor55.Appearance = appearance109;
			this.ultraNumericEditor55.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv13"));
			this.ultraNumericEditor55.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor55.FormatString = "#,##0.00";
			this.ultraNumericEditor55.Location = new System.Drawing.Point(592, 360);
			this.ultraNumericEditor55.Name = "ultraNumericEditor55";
			this.ultraNumericEditor55.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor55.PromptChar = ' ';
			this.ultraNumericEditor55.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor55.TabIndex = 95;
			this.ultraNumericEditor55.Validated += new System.EventHandler(this.ultraNumericEditor55_Validated);
			this.ultraNumericEditor55.ValueChanged += new System.EventHandler(this.ultraNumericEditor55_ValueChanged);
			this.ultraNumericEditor55.Enter += new System.EventHandler(this.ultraNumericEditor55_Enter);
			// 
			// ultraNumericEditor56
			// 
			appearance110.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor56.Appearance = appearance110;
			this.ultraNumericEditor56.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb13"));
			this.ultraNumericEditor56.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor56.FormatString = "";
			this.ultraNumericEditor56.Location = new System.Drawing.Point(512, 360);
			this.ultraNumericEditor56.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor56.Name = "ultraNumericEditor56";
			this.ultraNumericEditor56.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor56.PromptChar = ' ';
			this.ultraNumericEditor56.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor56.TabIndex = 94;
			this.ultraNumericEditor56.ValueChanged += new System.EventHandler(this.ultraNumericEditor56_ValueChanged);
			// 
			// ultraCombo13
			// 
			appearance111.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo13.Appearance = appearance111;
			this.ultraCombo13.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo13.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo13.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt13"));
			this.ultraCombo13.DataSource = this.cdvFormaPago;
			ultraGridBand14.AddButtonCaption = "MarceloBodega";
			this.ultraCombo13.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.ultraCombo13.DisplayMember = "FormaPago";
			this.ultraCombo13.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo13.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo13.Location = new System.Drawing.Point(384, 360);
			this.ultraCombo13.Name = "ultraCombo13";
			this.ultraCombo13.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo13.TabIndex = 93;
			this.ultraCombo13.ValueMember = "idFormaPago";
			this.ultraCombo13.Validated += new System.EventHandler(this.ultraCombo13_Validated);
			// 
			// ultraNumericEditor57
			// 
			appearance112.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor57.Appearance = appearance112;
			this.ultraNumericEditor57.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv14"));
			this.ultraNumericEditor57.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor57.FormatString = "#,##0.00";
			this.ultraNumericEditor57.Location = new System.Drawing.Point(592, 384);
			this.ultraNumericEditor57.Name = "ultraNumericEditor57";
			this.ultraNumericEditor57.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor57.PromptChar = ' ';
			this.ultraNumericEditor57.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor57.TabIndex = 98;
			this.ultraNumericEditor57.Validated += new System.EventHandler(this.ultraNumericEditor57_Validated);
			this.ultraNumericEditor57.Enter += new System.EventHandler(this.ultraNumericEditor57_Enter);
			// 
			// ultraNumericEditor58
			// 
			appearance113.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor58.Appearance = appearance113;
			this.ultraNumericEditor58.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb14"));
			this.ultraNumericEditor58.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor58.FormatString = "";
			this.ultraNumericEditor58.Location = new System.Drawing.Point(512, 384);
			this.ultraNumericEditor58.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor58.Name = "ultraNumericEditor58";
			this.ultraNumericEditor58.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor58.PromptChar = ' ';
			this.ultraNumericEditor58.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor58.TabIndex = 97;
			// 
			// ultraCombo14
			// 
			appearance114.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo14.Appearance = appearance114;
			this.ultraCombo14.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo14.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo14.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt14"));
			this.ultraCombo14.DataSource = this.cdvFormaPago;
			ultraGridBand15.AddButtonCaption = "MarceloBodega";
			this.ultraCombo14.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.ultraCombo14.DisplayMember = "FormaPago";
			this.ultraCombo14.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo14.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo14.Location = new System.Drawing.Point(384, 384);
			this.ultraCombo14.Name = "ultraCombo14";
			this.ultraCombo14.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo14.TabIndex = 96;
			this.ultraCombo14.ValueMember = "idFormaPago";
			this.ultraCombo14.Validated += new System.EventHandler(this.ultraCombo14_Validated);
			this.ultraCombo14.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo14_InitializeLayout);
			// 
			// ultraNumericEditor59
			// 
			appearance115.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor59.Appearance = appearance115;
			this.ultraNumericEditor59.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv15"));
			this.ultraNumericEditor59.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor59.FormatString = "#,##0.00";
			this.ultraNumericEditor59.Location = new System.Drawing.Point(592, 408);
			this.ultraNumericEditor59.Name = "ultraNumericEditor59";
			this.ultraNumericEditor59.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor59.PromptChar = ' ';
			this.ultraNumericEditor59.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor59.TabIndex = 101;
			this.ultraNumericEditor59.Validated += new System.EventHandler(this.ultraNumericEditor59_Validated);
			this.ultraNumericEditor59.ValueChanged += new System.EventHandler(this.ultraNumericEditor59_ValueChanged);
			this.ultraNumericEditor59.Enter += new System.EventHandler(this.ultraNumericEditor59_Enter);
			// 
			// ultraNumericEditor60
			// 
			appearance116.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor60.Appearance = appearance116;
			this.ultraNumericEditor60.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb15"));
			this.ultraNumericEditor60.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor60.FormatString = "";
			this.ultraNumericEditor60.Location = new System.Drawing.Point(512, 408);
			this.ultraNumericEditor60.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor60.Name = "ultraNumericEditor60";
			this.ultraNumericEditor60.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor60.PromptChar = ' ';
			this.ultraNumericEditor60.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor60.TabIndex = 100;
			// 
			// ultraCombo15
			// 
			appearance117.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo15.Appearance = appearance117;
			this.ultraCombo15.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo15.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo15.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt15"));
			this.ultraCombo15.DataSource = this.cdvFormaPago;
			ultraGridBand16.AddButtonCaption = "MarceloBodega";
			this.ultraCombo15.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.ultraCombo15.DisplayMember = "FormaPago";
			this.ultraCombo15.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo15.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo15.Location = new System.Drawing.Point(384, 408);
			this.ultraCombo15.Name = "ultraCombo15";
			this.ultraCombo15.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo15.TabIndex = 99;
			this.ultraCombo15.ValueMember = "idFormaPago";
			this.ultraCombo15.Validated += new System.EventHandler(this.ultraCombo15_Validated);
			// 
			// ultraNumericEditor61
			// 
			appearance118.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor61.Appearance = appearance118;
			this.ultraNumericEditor61.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv16"));
			this.ultraNumericEditor61.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor61.FormatString = "#,##0.00";
			this.ultraNumericEditor61.Location = new System.Drawing.Point(592, 432);
			this.ultraNumericEditor61.Name = "ultraNumericEditor61";
			this.ultraNumericEditor61.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor61.PromptChar = ' ';
			this.ultraNumericEditor61.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor61.TabIndex = 104;
			this.ultraNumericEditor61.Validated += new System.EventHandler(this.ultraNumericEditor61_Validated);
			this.ultraNumericEditor61.Enter += new System.EventHandler(this.ultraNumericEditor61_Enter);
			// 
			// ultraNumericEditor62
			// 
			appearance119.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor62.Appearance = appearance119;
			this.ultraNumericEditor62.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb16"));
			this.ultraNumericEditor62.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor62.FormatString = "";
			this.ultraNumericEditor62.Location = new System.Drawing.Point(512, 432);
			this.ultraNumericEditor62.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor62.Name = "ultraNumericEditor62";
			this.ultraNumericEditor62.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor62.PromptChar = ' ';
			this.ultraNumericEditor62.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor62.TabIndex = 103;
			// 
			// ultraCombo16
			// 
			appearance120.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo16.Appearance = appearance120;
			this.ultraCombo16.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo16.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo16.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt16"));
			this.ultraCombo16.DataSource = this.cdvFormaPago;
			ultraGridBand17.AddButtonCaption = "MarceloBodega";
			this.ultraCombo16.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.ultraCombo16.DisplayMember = "FormaPago";
			this.ultraCombo16.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo16.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo16.Location = new System.Drawing.Point(384, 432);
			this.ultraCombo16.Name = "ultraCombo16";
			this.ultraCombo16.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo16.TabIndex = 102;
			this.ultraCombo16.ValueMember = "idFormaPago";
			this.ultraCombo16.Validated += new System.EventHandler(this.ultraCombo16_Validated);
			this.ultraCombo16.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo16_InitializeLayout);
			// 
			// ultraNumericEditor63
			// 
			appearance121.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor63.Appearance = appearance121;
			this.ultraNumericEditor63.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv17"));
			this.ultraNumericEditor63.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor63.FormatString = "#,##0.00";
			this.ultraNumericEditor63.Location = new System.Drawing.Point(592, 456);
			this.ultraNumericEditor63.Name = "ultraNumericEditor63";
			this.ultraNumericEditor63.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor63.PromptChar = ' ';
			this.ultraNumericEditor63.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor63.TabIndex = 107;
			this.ultraNumericEditor63.Validated += new System.EventHandler(this.ultraNumericEditor63_Validated);
			this.ultraNumericEditor63.Enter += new System.EventHandler(this.ultraNumericEditor63_Enter);
			// 
			// ultraNumericEditor64
			// 
			appearance122.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor64.Appearance = appearance122;
			this.ultraNumericEditor64.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb17"));
			this.ultraNumericEditor64.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor64.FormatString = "";
			this.ultraNumericEditor64.Location = new System.Drawing.Point(512, 456);
			this.ultraNumericEditor64.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor64.Name = "ultraNumericEditor64";
			this.ultraNumericEditor64.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor64.PromptChar = ' ';
			this.ultraNumericEditor64.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor64.TabIndex = 106;
			this.ultraNumericEditor64.ValueChanged += new System.EventHandler(this.ultraNumericEditor64_ValueChanged_1);
			// 
			// ultraCombo17
			// 
			appearance123.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo17.Appearance = appearance123;
			this.ultraCombo17.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo17.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo17.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt17"));
			this.ultraCombo17.DataSource = this.cdvFormaPago;
			ultraGridBand18.AddButtonCaption = "MarceloBodega";
			this.ultraCombo17.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.ultraCombo17.DisplayMember = "FormaPago";
			this.ultraCombo17.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo17.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo17.Location = new System.Drawing.Point(384, 456);
			this.ultraCombo17.Name = "ultraCombo17";
			this.ultraCombo17.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo17.TabIndex = 105;
			this.ultraCombo17.ValueMember = "idFormaPago";
			this.ultraCombo17.Validated += new System.EventHandler(this.ultraCombo17_Validated);
			this.ultraCombo17.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo17_InitializeLayout);
			// 
			// ultraNumericEditor65
			// 
			appearance124.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor65.Appearance = appearance124;
			this.ultraNumericEditor65.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv18"));
			this.ultraNumericEditor65.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor65.FormatString = "#,##0.00";
			this.ultraNumericEditor65.Location = new System.Drawing.Point(592, 480);
			this.ultraNumericEditor65.Name = "ultraNumericEditor65";
			this.ultraNumericEditor65.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor65.PromptChar = ' ';
			this.ultraNumericEditor65.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor65.TabIndex = 110;
			this.ultraNumericEditor65.Validated += new System.EventHandler(this.ultraNumericEditor65_Validated);
			this.ultraNumericEditor65.ValueChanged += new System.EventHandler(this.ultraNumericEditor65_ValueChanged);
			this.ultraNumericEditor65.Enter += new System.EventHandler(this.ultraNumericEditor65_Enter);
			// 
			// ultraNumericEditor66
			// 
			appearance125.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor66.Appearance = appearance125;
			this.ultraNumericEditor66.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb18"));
			this.ultraNumericEditor66.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor66.FormatString = "";
			this.ultraNumericEditor66.Location = new System.Drawing.Point(512, 480);
			this.ultraNumericEditor66.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor66.Name = "ultraNumericEditor66";
			this.ultraNumericEditor66.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor66.PromptChar = ' ';
			this.ultraNumericEditor66.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor66.TabIndex = 109;
			// 
			// ultraCombo18
			// 
			appearance126.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo18.Appearance = appearance126;
			this.ultraCombo18.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo18.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo18.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt18"));
			this.ultraCombo18.DataSource = this.cdvFormaPago;
			ultraGridBand19.AddButtonCaption = "MarceloBodega";
			this.ultraCombo18.DisplayLayout.BandsSerializer.Add(ultraGridBand19);
			this.ultraCombo18.DisplayMember = "FormaPago";
			this.ultraCombo18.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo18.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo18.Location = new System.Drawing.Point(384, 480);
			this.ultraCombo18.Name = "ultraCombo18";
			this.ultraCombo18.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo18.TabIndex = 108;
			this.ultraCombo18.ValueMember = "idFormaPago";
			this.ultraCombo18.Validated += new System.EventHandler(this.ultraCombo18_Validated);
			// 
			// ultraNumericEditor67
			// 
			appearance127.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor67.Appearance = appearance127;
			this.ultraNumericEditor67.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv19"));
			this.ultraNumericEditor67.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor67.FormatString = "#,##0.00";
			this.ultraNumericEditor67.Location = new System.Drawing.Point(592, 504);
			this.ultraNumericEditor67.Name = "ultraNumericEditor67";
			this.ultraNumericEditor67.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor67.PromptChar = ' ';
			this.ultraNumericEditor67.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor67.TabIndex = 113;
			this.ultraNumericEditor67.Validated += new System.EventHandler(this.ultraNumericEditor67_Validated);
			this.ultraNumericEditor67.Enter += new System.EventHandler(this.ultraNumericEditor67_Enter);
			// 
			// ultraNumericEditor68
			// 
			appearance128.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor68.Appearance = appearance128;
			this.ultraNumericEditor68.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb19"));
			this.ultraNumericEditor68.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor68.FormatString = "";
			this.ultraNumericEditor68.Location = new System.Drawing.Point(512, 504);
			this.ultraNumericEditor68.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor68.Name = "ultraNumericEditor68";
			this.ultraNumericEditor68.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor68.PromptChar = ' ';
			this.ultraNumericEditor68.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor68.TabIndex = 112;
			// 
			// ultraCombo19
			// 
			appearance129.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo19.Appearance = appearance129;
			this.ultraCombo19.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo19.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo19.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt19"));
			this.ultraCombo19.DataSource = this.cdvFormaPago;
			ultraGridBand20.AddButtonCaption = "MarceloBodega";
			this.ultraCombo19.DisplayLayout.BandsSerializer.Add(ultraGridBand20);
			this.ultraCombo19.DisplayMember = "FormaPago";
			this.ultraCombo19.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo19.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo19.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo19.Location = new System.Drawing.Point(384, 504);
			this.ultraCombo19.Name = "ultraCombo19";
			this.ultraCombo19.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo19.TabIndex = 111;
			this.ultraCombo19.ValueMember = "idFormaPago";
			this.ultraCombo19.Validated += new System.EventHandler(this.ultraCombo19_Validated);
			this.ultraCombo19.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo19_InitializeLayout);
			// 
			// ultraNumericEditor69
			// 
			appearance130.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor69.Appearance = appearance130;
			this.ultraNumericEditor69.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mv20"));
			this.ultraNumericEditor69.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor69.FormatString = "#,##0.00";
			this.ultraNumericEditor69.Location = new System.Drawing.Point(592, 528);
			this.ultraNumericEditor69.Name = "ultraNumericEditor69";
			this.ultraNumericEditor69.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor69.PromptChar = ' ';
			this.ultraNumericEditor69.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor69.TabIndex = 116;
			this.ultraNumericEditor69.Validated += new System.EventHandler(this.ultraNumericEditor69_Validated);
			this.ultraNumericEditor69.Enter += new System.EventHandler(this.ultraNumericEditor69_Enter);
			// 
			// ultraNumericEditor70
			// 
			appearance131.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor70.Appearance = appearance131;
			this.ultraNumericEditor70.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mb20"));
			this.ultraNumericEditor70.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor70.FormatString = "";
			this.ultraNumericEditor70.Location = new System.Drawing.Point(512, 528);
			this.ultraNumericEditor70.MaskInput = "nnnnnnnnnnn";
			this.ultraNumericEditor70.Name = "ultraNumericEditor70";
			this.ultraNumericEditor70.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor70.PromptChar = ' ';
			this.ultraNumericEditor70.Size = new System.Drawing.Size(72, 21);
			this.ultraNumericEditor70.TabIndex = 115;
			this.ultraNumericEditor70.ValueChanged += new System.EventHandler(this.ultraNumericEditor70_ValueChanged);
			// 
			// ultraCombo20
			// 
			appearance132.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo20.Appearance = appearance132;
			this.ultraCombo20.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo20.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo20.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "mt20"));
			this.ultraCombo20.DataSource = this.cdvFormaPago;
			ultraGridBand21.AddButtonCaption = "MarceloBodega";
			this.ultraCombo20.DisplayLayout.BandsSerializer.Add(ultraGridBand21);
			this.ultraCombo20.DisplayMember = "FormaPago";
			this.ultraCombo20.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo20.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo20.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo20.Location = new System.Drawing.Point(384, 528);
			this.ultraCombo20.Name = "ultraCombo20";
			this.ultraCombo20.Size = new System.Drawing.Size(128, 21);
			this.ultraCombo20.TabIndex = 114;
			this.ultraCombo20.ValueMember = "idFormaPago";
			this.ultraCombo20.Validated += new System.EventHandler(this.ultraCombo20_Validated);
			// 
			// txtNombre1
			// 
			appearance133.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre1.Appearance = appearance133;
			this.txtNombre1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cet1"));
			this.txtNombre1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre1.Location = new System.Drawing.Point(16, 184);
			this.txtNombre1.Name = "txtNombre1";
			this.txtNombre1.Size = new System.Drawing.Size(224, 21);
			this.txtNombre1.TabIndex = 12;
			this.txtNombre1.Enter += new System.EventHandler(this.txtNombre1_Enter);
			// 
			// txtNombre2
			// 
			appearance134.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre2.Appearance = appearance134;
			this.txtNombre2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cet2"));
			this.txtNombre2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre2.Location = new System.Drawing.Point(16, 208);
			this.txtNombre2.Name = "txtNombre2";
			this.txtNombre2.Size = new System.Drawing.Size(224, 21);
			this.txtNombre2.TabIndex = 15;
			this.txtNombre2.Enter += new System.EventHandler(this.txtNombre2_Enter);
			// 
			// txtNombre3
			// 
			appearance135.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre3.Appearance = appearance135;
			this.txtNombre3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cet3"));
			this.txtNombre3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre3.Location = new System.Drawing.Point(16, 232);
			this.txtNombre3.Name = "txtNombre3";
			this.txtNombre3.Size = new System.Drawing.Size(224, 21);
			this.txtNombre3.TabIndex = 18;
			this.txtNombre3.Enter += new System.EventHandler(this.txtNombre3_Enter);
			// 
			// txtNombre4
			// 
			appearance136.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre4.Appearance = appearance136;
			this.txtNombre4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cet4"));
			this.txtNombre4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre4.Location = new System.Drawing.Point(16, 256);
			this.txtNombre4.Name = "txtNombre4";
			this.txtNombre4.Size = new System.Drawing.Size(224, 21);
			this.txtNombre4.TabIndex = 21;
			this.txtNombre4.Enter += new System.EventHandler(this.txtNombre4_Enter);
			// 
			// cefact1
			// 
			appearance137.ForeColorDisabled = System.Drawing.Color.Black;
			this.cefact1.Appearance = appearance137;
			this.cefact1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cef1"));
			this.cefact1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cefact1.FormatString = "";
			this.cefact1.Location = new System.Drawing.Point(248, 184);
			this.cefact1.MaskInput = "nnnnnnn";
			this.cefact1.Name = "cefact1";
			this.cefact1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.cefact1.PromptChar = ' ';
			this.cefact1.Size = new System.Drawing.Size(56, 21);
			this.cefact1.TabIndex = 13;
			this.cefact1.Enter += new System.EventHandler(this.cefact1_Enter);
			// 
			// cefact2
			// 
			appearance138.ForeColorDisabled = System.Drawing.Color.Black;
			this.cefact2.Appearance = appearance138;
			this.cefact2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cef2"));
			this.cefact2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cefact2.FormatString = "";
			this.cefact2.Location = new System.Drawing.Point(248, 208);
			this.cefact2.MaskInput = "nnnnnnn";
			this.cefact2.Name = "cefact2";
			this.cefact2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.cefact2.PromptChar = ' ';
			this.cefact2.Size = new System.Drawing.Size(56, 21);
			this.cefact2.TabIndex = 16;
			this.cefact2.Enter += new System.EventHandler(this.cefact2_Enter);
			// 
			// cefact3
			// 
			appearance139.ForeColorDisabled = System.Drawing.Color.Black;
			this.cefact3.Appearance = appearance139;
			this.cefact3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cef3"));
			this.cefact3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cefact3.FormatString = "";
			this.cefact3.Location = new System.Drawing.Point(248, 232);
			this.cefact3.MaskInput = "nnnnnnn";
			this.cefact3.Name = "cefact3";
			this.cefact3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.cefact3.PromptChar = ' ';
			this.cefact3.Size = new System.Drawing.Size(56, 21);
			this.cefact3.TabIndex = 19;
			this.cefact3.Enter += new System.EventHandler(this.cefact3_Enter);
			// 
			// cefact4
			// 
			appearance140.ForeColorDisabled = System.Drawing.Color.Black;
			this.cefact4.Appearance = appearance140;
			this.cefact4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cef4"));
			this.cefact4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cefact4.FormatString = "";
			this.cefact4.Location = new System.Drawing.Point(248, 256);
			this.cefact4.MaskInput = "nnnnnnn";
			this.cefact4.Name = "cefact4";
			this.cefact4.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.cefact4.PromptChar = ' ';
			this.cefact4.Size = new System.Drawing.Size(56, 21);
			this.cefact4.TabIndex = 22;
			this.cefact4.Enter += new System.EventHandler(this.cefact4_Enter);
			// 
			// cet1
			// 
			appearance141.ForeColorDisabled = System.Drawing.Color.Black;
			this.cet1.Appearance = appearance141;
			this.cet1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cev1"));
			this.cet1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cet1.FormatString = "#,##0.00";
			this.cet1.Location = new System.Drawing.Point(312, 184);
			this.cet1.Name = "cet1";
			this.cet1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.cet1.PromptChar = ' ';
			this.cet1.Size = new System.Drawing.Size(56, 21);
			this.cet1.TabIndex = 14;
			this.cet1.Validated += new System.EventHandler(this.cet1_Validated);
			this.cet1.Enter += new System.EventHandler(this.cet1_Enter);
			// 
			// cet2
			// 
			appearance142.ForeColorDisabled = System.Drawing.Color.Black;
			this.cet2.Appearance = appearance142;
			this.cet2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cev2"));
			this.cet2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cet2.FormatString = "#,##0.00";
			this.cet2.Location = new System.Drawing.Point(312, 208);
			this.cet2.Name = "cet2";
			this.cet2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.cet2.PromptChar = ' ';
			this.cet2.Size = new System.Drawing.Size(56, 21);
			this.cet2.TabIndex = 17;
			this.cet2.Validated += new System.EventHandler(this.cet2_Validated);
			this.cet2.Enter += new System.EventHandler(this.cet2_Enter);
			// 
			// cet3
			// 
			appearance143.ForeColorDisabled = System.Drawing.Color.Black;
			this.cet3.Appearance = appearance143;
			this.cet3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cev3"));
			this.cet3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cet3.FormatString = "#,##0.00";
			this.cet3.Location = new System.Drawing.Point(312, 232);
			this.cet3.Name = "cet3";
			this.cet3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.cet3.PromptChar = ' ';
			this.cet3.Size = new System.Drawing.Size(56, 21);
			this.cet3.TabIndex = 20;
			this.cet3.Validated += new System.EventHandler(this.cet3_Validated);
			this.cet3.Enter += new System.EventHandler(this.cet3_Enter);
			// 
			// cet4
			// 
			appearance144.ForeColorDisabled = System.Drawing.Color.Black;
			this.cet4.Appearance = appearance144;
			this.cet4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cev4"));
			this.cet4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cet4.FormatString = "#,##0.00";
			this.cet4.Location = new System.Drawing.Point(312, 256);
			this.cet4.Name = "cet4";
			this.cet4.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.cet4.PromptChar = ' ';
			this.cet4.Size = new System.Drawing.Size(56, 21);
			this.cet4.TabIndex = 23;
			this.cet4.Validated += new System.EventHandler(this.cet4_Validated);
			this.cet4.Enter += new System.EventHandler(this.cet4_Enter);
			// 
			// cet5
			// 
			appearance145.ForeColorDisabled = System.Drawing.Color.Black;
			this.cet5.Appearance = appearance145;
			this.cet5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cev5"));
			this.cet5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cet5.FormatString = "#,##0.00";
			this.cet5.Location = new System.Drawing.Point(312, 280);
			this.cet5.Name = "cet5";
			this.cet5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.cet5.PromptChar = ' ';
			this.cet5.Size = new System.Drawing.Size(56, 21);
			this.cet5.TabIndex = 26;
			this.cet5.Validated += new System.EventHandler(this.cet5_Validated);
			this.cet5.Enter += new System.EventHandler(this.cet5_Enter);
			// 
			// cefact5
			// 
			appearance146.ForeColorDisabled = System.Drawing.Color.Black;
			this.cefact5.Appearance = appearance146;
			this.cefact5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cef5"));
			this.cefact5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cefact5.FormatString = "";
			this.cefact5.Location = new System.Drawing.Point(248, 280);
			this.cefact5.MaskInput = "nnnnnnn";
			this.cefact5.Name = "cefact5";
			this.cefact5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.cefact5.PromptChar = ' ';
			this.cefact5.Size = new System.Drawing.Size(56, 21);
			this.cefact5.TabIndex = 25;
			this.cefact5.Enter += new System.EventHandler(this.cefact5_Enter);
			// 
			// txtNombre5
			// 
			appearance147.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre5.Appearance = appearance147;
			this.txtNombre5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "cet5"));
			this.txtNombre5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre5.Location = new System.Drawing.Point(16, 280);
			this.txtNombre5.Name = "txtNombre5";
			this.txtNombre5.Size = new System.Drawing.Size(224, 21);
			this.txtNombre5.TabIndex = 24;
			this.txtNombre5.Enter += new System.EventHandler(this.txtNombre5_Enter);
			// 
			// btTraer
			// 
			this.btTraer.Location = new System.Drawing.Point(840, 552);
			this.btTraer.Name = "btTraer";
			this.btTraer.Size = new System.Drawing.Size(80, 23);
			this.btTraer.TabIndex = 272;
			this.btTraer.Text = "Traer Venta Diaria";
			this.btTraer.Click += new System.EventHandler(this.button1_Click);
			// 
			// label72
			// 
			this.label72.AutoSize = true;
			this.label72.Location = new System.Drawing.Point(40, 168);
			this.label72.Name = "label72";
			this.label72.Size = new System.Drawing.Size(193, 16);
			this.label72.TabIndex = 273;
			this.label72.Text = "EMPLEADO - CREDITO PERSONAL";
			// 
			// label82
			// 
			this.label82.AutoSize = true;
			this.label82.Location = new System.Drawing.Point(248, 168);
			this.label82.Name = "label82";
			this.label82.Size = new System.Drawing.Size(51, 16);
			this.label82.TabIndex = 274;
			this.label82.Text = "N. FACT.";
			// 
			// label84
			// 
			this.label84.AutoSize = true;
			this.label84.Location = new System.Drawing.Point(320, 168);
			this.label84.Name = "label84";
			this.label84.Size = new System.Drawing.Size(43, 16);
			this.label84.TabIndex = 275;
			this.label84.Text = "VALOR";
			// 
			// ultraTextEditor1
			// 
			appearance148.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor1.Appearance = appearance148;
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ncn1"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(736, 368);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(128, 21);
			this.ultraTextEditor1.TabIndex = 276;
			// 
			// ultraTextEditor2
			// 
			appearance149.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor2.Appearance = appearance149;
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ncn2"));
			this.ultraTextEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor2.Location = new System.Drawing.Point(736, 392);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(128, 21);
			this.ultraTextEditor2.TabIndex = 278;
			// 
			// ultraTextEditor3
			// 
			appearance150.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor3.Appearance = appearance150;
			this.ultraTextEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ncn3"));
			this.ultraTextEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor3.Location = new System.Drawing.Point(736, 416);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(128, 21);
			this.ultraTextEditor3.TabIndex = 280;
			// 
			// ultraTextEditor4
			// 
			appearance151.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor4.Appearance = appearance151;
			this.ultraTextEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ncn4"));
			this.ultraTextEditor4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor4.Location = new System.Drawing.Point(736, 440);
			this.ultraTextEditor4.Name = "ultraTextEditor4";
			this.ultraTextEditor4.Size = new System.Drawing.Size(128, 21);
			this.ultraTextEditor4.TabIndex = 282;
			// 
			// ultraTextEditor5
			// 
			appearance152.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraTextEditor5.Appearance = appearance152;
			this.ultraTextEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ncn5"));
			this.ultraTextEditor5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor5.Location = new System.Drawing.Point(736, 464);
			this.ultraTextEditor5.Name = "ultraTextEditor5";
			this.ultraTextEditor5.Size = new System.Drawing.Size(128, 21);
			this.ultraTextEditor5.TabIndex = 284;
			// 
			// ultraNumericEditor71
			// 
			appearance153.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor71.Appearance = appearance153;
			this.ultraNumericEditor71.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ncv1"));
			this.ultraNumericEditor71.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor71.FormatString = "#,##0.00";
			this.ultraNumericEditor71.Location = new System.Drawing.Point(880, 368);
			this.ultraNumericEditor71.Name = "ultraNumericEditor71";
			this.ultraNumericEditor71.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor71.PromptChar = ' ';
			this.ultraNumericEditor71.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor71.TabIndex = 277;
			// 
			// ultraNumericEditor72
			// 
			appearance154.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor72.Appearance = appearance154;
			this.ultraNumericEditor72.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ncv2"));
			this.ultraNumericEditor72.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor72.FormatString = "#,##0.00";
			this.ultraNumericEditor72.Location = new System.Drawing.Point(880, 392);
			this.ultraNumericEditor72.Name = "ultraNumericEditor72";
			this.ultraNumericEditor72.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor72.PromptChar = ' ';
			this.ultraNumericEditor72.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor72.TabIndex = 279;
			// 
			// ultraNumericEditor73
			// 
			appearance155.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor73.Appearance = appearance155;
			this.ultraNumericEditor73.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ncv3"));
			this.ultraNumericEditor73.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor73.FormatString = "#,##0.00";
			this.ultraNumericEditor73.Location = new System.Drawing.Point(880, 416);
			this.ultraNumericEditor73.Name = "ultraNumericEditor73";
			this.ultraNumericEditor73.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor73.PromptChar = ' ';
			this.ultraNumericEditor73.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor73.TabIndex = 281;
			// 
			// ultraNumericEditor74
			// 
			appearance156.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor74.Appearance = appearance156;
			this.ultraNumericEditor74.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ncv4"));
			this.ultraNumericEditor74.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor74.FormatString = "#,##0.00";
			this.ultraNumericEditor74.Location = new System.Drawing.Point(880, 440);
			this.ultraNumericEditor74.Name = "ultraNumericEditor74";
			this.ultraNumericEditor74.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor74.PromptChar = ' ';
			this.ultraNumericEditor74.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor74.TabIndex = 283;
			// 
			// ultraNumericEditor75
			// 
			appearance157.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor75.Appearance = appearance157;
			this.ultraNumericEditor75.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvMarcelo, "ncv5"));
			this.ultraNumericEditor75.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor75.FormatString = "#,##0.00";
			this.ultraNumericEditor75.Location = new System.Drawing.Point(880, 464);
			this.ultraNumericEditor75.Name = "ultraNumericEditor75";
			this.ultraNumericEditor75.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor75.PromptChar = ' ';
			this.ultraNumericEditor75.Size = new System.Drawing.Size(56, 21);
			this.ultraNumericEditor75.TabIndex = 285;
			// 
			// label85
			// 
			this.label85.AutoSize = true;
			this.label85.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label85.Location = new System.Drawing.Point(728, 344);
			this.label85.Name = "label85";
			this.label85.Size = new System.Drawing.Size(119, 16);
			this.label85.TabIndex = 286;
			this.label85.Text = "NOTAS DE CREDITO";
			// 
			// MarceloGeneralVentas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1008, 581);
			this.Controls.Add(this.label85);
			this.Controls.Add(this.ultraNumericEditor75);
			this.Controls.Add(this.ultraNumericEditor74);
			this.Controls.Add(this.ultraNumericEditor73);
			this.Controls.Add(this.ultraNumericEditor72);
			this.Controls.Add(this.ultraNumericEditor71);
			this.Controls.Add(this.ultraTextEditor5);
			this.Controls.Add(this.ultraTextEditor4);
			this.Controls.Add(this.ultraTextEditor3);
			this.Controls.Add(this.ultraTextEditor2);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.label84);
			this.Controls.Add(this.label82);
			this.Controls.Add(this.label72);
			this.Controls.Add(this.btTraer);
			this.Controls.Add(this.cet5);
			this.Controls.Add(this.cefact5);
			this.Controls.Add(this.txtNombre5);
			this.Controls.Add(this.cet4);
			this.Controls.Add(this.cet3);
			this.Controls.Add(this.cet2);
			this.Controls.Add(this.cet1);
			this.Controls.Add(this.cefact4);
			this.Controls.Add(this.cefact3);
			this.Controls.Add(this.cefact2);
			this.Controls.Add(this.cefact1);
			this.Controls.Add(this.txtNombre4);
			this.Controls.Add(this.txtNombre3);
			this.Controls.Add(this.txtNombre2);
			this.Controls.Add(this.txtNombre1);
			this.Controls.Add(this.ultraNumericEditor69);
			this.Controls.Add(this.ultraNumericEditor70);
			this.Controls.Add(this.ultraCombo20);
			this.Controls.Add(this.ultraNumericEditor67);
			this.Controls.Add(this.ultraNumericEditor68);
			this.Controls.Add(this.ultraCombo19);
			this.Controls.Add(this.ultraNumericEditor65);
			this.Controls.Add(this.ultraNumericEditor66);
			this.Controls.Add(this.ultraCombo18);
			this.Controls.Add(this.ultraNumericEditor63);
			this.Controls.Add(this.ultraNumericEditor64);
			this.Controls.Add(this.ultraCombo17);
			this.Controls.Add(this.ultraNumericEditor61);
			this.Controls.Add(this.ultraNumericEditor62);
			this.Controls.Add(this.ultraCombo16);
			this.Controls.Add(this.ultraNumericEditor59);
			this.Controls.Add(this.ultraNumericEditor60);
			this.Controls.Add(this.ultraCombo15);
			this.Controls.Add(this.ultraNumericEditor57);
			this.Controls.Add(this.ultraNumericEditor58);
			this.Controls.Add(this.ultraCombo14);
			this.Controls.Add(this.ultraNumericEditor55);
			this.Controls.Add(this.ultraNumericEditor56);
			this.Controls.Add(this.ultraCombo13);
			this.Controls.Add(this.ultraNumericEditor44);
			this.Controls.Add(this.ultraNumericEditor54);
			this.Controls.Add(this.ultraCombo12);
			this.Controls.Add(this.ultraNumericEditor43);
			this.Controls.Add(this.ultraNumericEditor32);
			this.Controls.Add(this.ultraCombo11);
			this.Controls.Add(this.ultraCombo10);
			this.Controls.Add(this.ultraCombo9);
			this.Controls.Add(this.ultraCombo8);
			this.Controls.Add(this.ultraCombo7);
			this.Controls.Add(this.ultraCombo6);
			this.Controls.Add(this.ultraCombo5);
			this.Controls.Add(this.ultraCombo4);
			this.Controls.Add(this.ultraCombo3);
			this.Controls.Add(this.ultraCombo2);
			this.Controls.Add(this.ultraCombo1);
			this.Controls.Add(this.ultraNumericEditor21);
			this.Controls.Add(this.ultraNumericEditor2);
			this.Controls.Add(this.txtTerminaNC);
			this.Controls.Add(this.txtIniciaNC);
			this.Controls.Add(this.label97);
			this.Controls.Add(this.label98);
			this.Controls.Add(this.label96);
			this.Controls.Add(this.label95);
			this.Controls.Add(this.label94);
			this.Controls.Add(this.label93);
			this.Controls.Add(this.label92);
			this.Controls.Add(this.label91);
			this.Controls.Add(this.label88);
			this.Controls.Add(this.label87);
			this.Controls.Add(this.label86);
			this.Controls.Add(this.label83);
			this.Controls.Add(this.label81);
			this.Controls.Add(this.label80);
			this.Controls.Add(this.label79);
			this.Controls.Add(this.label78);
			this.Controls.Add(this.label77);
			this.Controls.Add(this.label76);
			this.Controls.Add(this.label75);
			this.Controls.Add(this.label74);
			this.Controls.Add(this.label73);
			this.Controls.Add(this.label71);
			this.Controls.Add(this.label70);
			this.Controls.Add(this.label52);
			this.Controls.Add(this.label69);
			this.Controls.Add(this.label68);
			this.Controls.Add(this.label67);
			this.Controls.Add(this.label66);
			this.Controls.Add(this.label65);
			this.Controls.Add(this.label64);
			this.Controls.Add(this.label63);
			this.Controls.Add(this.label62);
			this.Controls.Add(this.label61);
			this.Controls.Add(this.label60);
			this.Controls.Add(this.label59);
			this.Controls.Add(this.label58);
			this.Controls.Add(this.label57);
			this.Controls.Add(this.label56);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.label54);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.label51);
			this.Controls.Add(this.label50);
			this.Controls.Add(this.label49);
			this.Controls.Add(this.label48);
			this.Controls.Add(this.label47);
			this.Controls.Add(this.label46);
			this.Controls.Add(this.label42);
			this.Controls.Add(this.label43);
			this.Controls.Add(this.label44);
			this.Controls.Add(this.label45);
			this.Controls.Add(this.label41);
			this.Controls.Add(this.label40);
			this.Controls.Add(this.label39);
			this.Controls.Add(this.label38);
			this.Controls.Add(this.label37);
			this.Controls.Add(this.txtCredEmpresa);
			this.Controls.Add(this.txtRetencion);
			this.Controls.Add(this.txtPromo79);
			this.Controls.Add(this.txtPromo35);
			this.Controls.Add(this.txtTerminaF);
			this.Controls.Add(this.txtIniciaF);
			this.Controls.Add(this.ultraNumericEditor45);
			this.Controls.Add(this.ultraNumericEditor46);
			this.Controls.Add(this.ultraNumericEditor47);
			this.Controls.Add(this.ultraNumericEditor48);
			this.Controls.Add(this.ultraNumericEditor49);
			this.Controls.Add(this.ultraNumericEditor50);
			this.Controls.Add(this.ultraNumericEditor51);
			this.Controls.Add(this.ultraNumericEditor52);
			this.Controls.Add(this.ultraNumericEditor53);
			this.Controls.Add(this.ultraNumericEditor34);
			this.Controls.Add(this.ultraNumericEditor35);
			this.Controls.Add(this.ultraNumericEditor36);
			this.Controls.Add(this.ultraNumericEditor37);
			this.Controls.Add(this.ultraNumericEditor38);
			this.Controls.Add(this.ultraNumericEditor39);
			this.Controls.Add(this.ultraNumericEditor40);
			this.Controls.Add(this.ultraNumericEditor41);
			this.Controls.Add(this.ultraNumericEditor42);
			this.Controls.Add(this.ultraNumericEditor23);
			this.Controls.Add(this.ultraNumericEditor24);
			this.Controls.Add(this.ultraNumericEditor25);
			this.Controls.Add(this.ultraNumericEditor26);
			this.Controls.Add(this.ultraNumericEditor27);
			this.Controls.Add(this.ultraNumericEditor28);
			this.Controls.Add(this.ultraNumericEditor29);
			this.Controls.Add(this.ultraNumericEditor30);
			this.Controls.Add(this.ultraNumericEditor31);
			this.Controls.Add(this.ultraNumericEditor33);
			this.Controls.Add(this.ultraNumericEditor12);
			this.Controls.Add(this.ultraNumericEditor13);
			this.Controls.Add(this.ultraNumericEditor14);
			this.Controls.Add(this.ultraNumericEditor15);
			this.Controls.Add(this.ultraNumericEditor16);
			this.Controls.Add(this.ultraNumericEditor17);
			this.Controls.Add(this.ultraNumericEditor18);
			this.Controls.Add(this.ultraNumericEditor19);
			this.Controls.Add(this.ultraNumericEditor20);
			this.Controls.Add(this.ultraNumericEditor22);
			this.Controls.Add(this.ultraNumericEditor11);
			this.Controls.Add(this.ultraNumericEditor10);
			this.Controls.Add(this.ultraNumericEditor9);
			this.Controls.Add(this.ultraNumericEditor8);
			this.Controls.Add(this.ultraNumericEditor7);
			this.Controls.Add(this.ultraNumericEditor6);
			this.Controls.Add(this.ultraNumericEditor5);
			this.Controls.Add(this.ultraNumericEditor4);
			this.Controls.Add(this.ultraNumericEditor3);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.btReportes);
			this.Controls.Add(this.btDeshacer);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.txtPromo25);
			this.Controls.Add(this.txtPromo20);
			this.Controls.Add(this.txtPromoOtro);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtNiña);
			this.Controls.Add(this.txtCartera);
			this.Controls.Add(this.txtCinturon);
			this.Controls.Add(this.txtGorra);
			this.Controls.Add(this.txtMochila);
			this.Controls.Add(this.txtMaleta);
			this.Controls.Add(this.txtCanguro);
			this.Controls.Add(this.txtBilletera);
			this.Controls.Add(this.txtPortazapatos);
			this.Controls.Add(this.txtBalon);
			this.Controls.Add(this.txtPromo5);
			this.Controls.Add(this.txtPromo10);
			this.Controls.Add(this.txtPromo15);
			this.Controls.Add(this.txtPromo29);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblfecha);
			this.Controls.Add(this.txtNiño);
			this.Controls.Add(this.txtMujer);
			this.Controls.Add(this.txtHombre);
			this.Controls.Add(this.txtUnidadesAntes);
			this.Controls.Add(this.txtValorAntes);
			this.Controls.Add(this.txtUnidadActual);
			this.Controls.Add(this.txtValorActual);
			this.Controls.Add(this.txtDeposito);
			this.Controls.Add(this.txtCredPersonal);
			this.Controls.Add(this.txtTotalTarjetas);
			this.Controls.Add(this.txtCheque);
			this.Controls.Add(this.txtGastos);
			this.Controls.Add(this.txtOtrasTarjetas);
			this.Controls.Add(this.txtCuotaFacil);
			this.Controls.Add(this.txtEfectivo);
			this.Controls.Add(this.barraDato1);
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "MarceloGeneralVentas";
			this.Text = "Marcelo General Ventas";
			this.Load += new System.EventHandler(this.MarceloGeneralVentas_Load);
			this.Enter += new System.EventHandler(this.MarceloGeneralVentas_Enter);
			((System.ComponentModel.ISupportInitialize)(this.cdsMarcelo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarcelo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaFacil)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtrasTarjetas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalTarjetas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorActual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUnidadActual)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorAntes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUnidadesAntes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMujer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNiño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMarceloBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo29)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo15)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBalon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPortazapatos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBilletera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCanguro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMaleta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMochila)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGorra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCinturon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCartera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNiña)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromoOtro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo25)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor15)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor16)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor17)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor18)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor19)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor22)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor23)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor24)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor25)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor26)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor27)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor28)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor29)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor30)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor31)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor33)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor34)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor38)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor39)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor40)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor41)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor42)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor45)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor49)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor50)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor51)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor52)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor53)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIniciaF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTerminaF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo35)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromo79)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTerminaNC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIniciaNC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor48)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor37)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor36)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor47)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor46)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor35)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor21)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor32)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor43)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor44)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor54)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor55)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor56)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor57)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor58)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor59)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor60)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo15)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor61)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor62)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo16)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor63)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor64)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo17)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor65)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor66)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo18)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor67)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor68)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo19)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor69)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor70)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cefact1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cefact2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cefact3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cefact4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cet1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cet2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cet3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cet4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cet5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cefact5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor71)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor72)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor73)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor74)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor75)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

    C1.Data.C1DataRow drMarcelo;
		private void MarceloGeneralVentas_Load(object sender, System.EventArgs e)
		{
			this.cmbBodega.Value=5; 
			this.dtFecha.Value = DateTime.Today; 
			DateTime dtCaducidad = DateTime.Today;

		}

		private void cdsMarcelo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMarcelo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
			  drMarcelo  = cdsMarcelo.TableViews["MarceloGeneralVentas"].Rows[0];
			}
			catch{};
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
		barraDato1.DeshacerRegistro();  
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
		barraDato1.EditarRegistro();  
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			barraDato1.GrabaRegistro();  
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
		
		}

		private void cdsMarceloBodega_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMarceloBodega.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}


		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			actualizar();

		}

		private void cmbBodega_RowSelected_1(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
		{
			actualizar();

		}
		private void actualizar()
		{
			DateTime dtFecha1 = (DateTime) dtFecha.Value;
			this.lblfecha.Text= dtFecha1.ToString("yyyyMMdd") + cmbBodega.Value;
			int id=int.Parse(lblfecha.Text);
			barraDato1.IdRegistro =  id;
			barraDato1.ProximoId(5);
			

			//txtTotalTarjetas.Value=(double)txtCuotaFacil.Value + (double)txtOtrasTarjetas.Value;
			txtTotalTarjetas.Value = (double) drMarcelo["CoutaFacil"]+(double) drMarcelo["Tarjeta"];		
			barraDato1.EditarRegistro();  
			desactivar();
	
		}

		private void suma2()
		{
				double cetotal=0;
				cetotal=(double)drMarcelo["cev1"]+(double)drMarcelo["cev2"]+(double)drMarcelo["cev3"]+
					(double)drMarcelo["cev4"]+(double)drMarcelo["cev5"];
				
				drMarcelo["Credito"]=cetotal;
				
				double mtotal=0;
				double mcf=0;
				mtotal=(double)drMarcelo["mv1"]+(double)drMarcelo["mv2"]
				+(double)drMarcelo["mv3"]+(double)drMarcelo["mv4"]+(double)drMarcelo["mv5"]+(double)drMarcelo["mv6"]
				+(double)drMarcelo["mv7"]+(double)drMarcelo["mv8"]+(double)drMarcelo["mv9"]+(double)drMarcelo["mv10"]
				+(double)drMarcelo["mv11"]+(double)drMarcelo["mv12"]+(double)drMarcelo["mv13"]+(double)drMarcelo["mv14"]
				+(double)drMarcelo["mv15"]+(double)drMarcelo["mv16"]+(double)drMarcelo["mv17"]+(double)drMarcelo["mv18"]
				+(double)drMarcelo["mv19"]+(double)drMarcelo["mv20"];
				
				int cuentacf=22;
				if((int)drMarcelo["mt1"]==cuentacf) 
					mcf=mcf+(double)drMarcelo["mv1"];
				if((int)drMarcelo["mt2"]==cuentacf) 
					mcf=mcf+(double)drMarcelo["mv2"];
			if((int)drMarcelo["mt3"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv3"];
			if((int)drMarcelo["mt4"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv4"];
			if((int)drMarcelo["mt5"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv5"];
			if((int)drMarcelo["mt6"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv6"];
			if((int)drMarcelo["mt7"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv7"];
			if((int)drMarcelo["mt8"]==cuentacf) 
				if((int)drMarcelo["mt1"]==cuentacf) 
					mcf=mcf+(double)drMarcelo["mv8"];
			if((int)drMarcelo["mt9"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv9"];
			if((int)drMarcelo["mt10"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv10"];
			if((int)drMarcelo["mt11"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv11"];
			if((int)drMarcelo["mt12"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv12"];
			if((int)drMarcelo["mt13"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv13"];
			if((int)drMarcelo["mt14"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv14"];
			if((int)drMarcelo["mt15"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv15"];
			if((int)drMarcelo["mt16"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv16"];
			if((int)drMarcelo["mt17"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv17"];
			if((int)drMarcelo["mt18"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv18"];
			if((int)drMarcelo["mt19"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv19"];
			if((int)drMarcelo["mt20"]==cuentacf) 
				mcf=mcf+(double)drMarcelo["mv20"];

  	mtotal=	mtotal-mcf;

     double tarjeta;
			tarjeta= (double)drMarcelo["v1"]+ (double)drMarcelo["v2"]+ (double)drMarcelo["v3"]
				+ (double)drMarcelo["v4"]+ (double)drMarcelo["v5"]+ (double)drMarcelo["v6"]
				+ (double)drMarcelo["v7"]+ (double)drMarcelo["v9"]+ (double)drMarcelo["v10"]
				+ mtotal;
		
			double cf;
			cf=(double)drMarcelo["v8"]+mcf;
      
			drMarcelo["Tarjeta"]=tarjeta;
			drMarcelo["coutafacil"]=cf;

			txtTotalTarjetas.Value = (double) drMarcelo["CoutaFacil"]+(double) drMarcelo["Tarjeta"];		
		      		
			double deposito=0;
					
			deposito= (double)drMarcelo["Efectivo"]  + (double)drMarcelo["Cheque"]	- (double)drMarcelo["Gastos"]
			- (double)drMarcelo["Retencion"] -	(double)drMarcelo["CreditoEmpresa"];
			drMarcelo["Deposito"]= deposito;
 
			double totalVentas=0;
			totalVentas = (double)drMarcelo["Efectivo"] 
			+ (double)drMarcelo["Cheque"]
			+	(double)drMarcelo["CoutaFacil"]
			+ (double)drMarcelo["Tarjeta"]
			+	(double)drMarcelo["Credito"];
			
			drMarcelo["ValorActual"] = totalVentas;
		}
		
		private void suma()
		{
				double suma=0;
			  suma=(double)drMarcelo["Hombre"]+(double)drMarcelo["Mujer"]+(double)drMarcelo["Niño"]
				+(double)drMarcelo["Niña"]+(double)drMarcelo["Cartera"]+(double)drMarcelo["Cinturon"]
				+(double)drMarcelo["Gorra"]+(double)drMarcelo["Mochila"]+(double)drMarcelo["Canguro"]
				+(double)drMarcelo["Billetera"]+(double)drMarcelo["Portazapatos"]+(double)drMarcelo["Bolso"]
				+(double)drMarcelo["Promocion5"]+(double)drMarcelo["Promocion10"]+(double)drMarcelo["Promocion15"]
				+(double)drMarcelo["Promocion20"]+(double)drMarcelo["Promocion25"]+(double)drMarcelo["Promocion29"]
		    +(double)drMarcelo["Promocion35"]+(double)drMarcelo["Promocion79"]+(double)drMarcelo["Maleta"]
				+(double)drMarcelo["Otro"];
				drMarcelo["UnidadesActual"]= suma;
			
			/*drMarcelo["t1"]=(double)drMarcelo["v1"];	
				drMarcelo["t2"]=(double)drMarcelo["v2"];		
				drMarcelo["t3"]=(double)drMarcelo["v3"] + (double)drMarcelo["vv3"];		
				drMarcelo["t4"]=(double)drMarcelo["v4"] + (double)drMarcelo["vv4"];		
				drMarcelo["t5"]=(double)drMarcelo["v5"] + (double)drMarcelo["vv5"];		
				drMarcelo["t6"]=(double)drMarcelo["v6"] + (double)drMarcelo["vv6"];		
				drMarcelo["t7"]=(double)drMarcelo["v7"] + (double)drMarcelo["vv7"];		
				drMarcelo["t8"]=(double)drMarcelo["v8"] + (double)drMarcelo["vv8"];		
				drMarcelo["t9"]=(double)drMarcelo["v9"] + (double)drMarcelo["vv9"];		
				drMarcelo["t10"]=(double)drMarcelo["v10"] + (double)drMarcelo["vv10"];
				drMarcelo["t11"]=(double)drMarcelo["v11"] + (double)drMarcelo["vv11"];


				double s1 = (double)drMarcelo["v1"] + (double)drMarcelo["v2"];
				

				drMarcelo["TotalTarjetas"]= (double)drMarcelo["t1"]+ (double)drMarcelo["t2"]+(double)drMarcelo["t3"]
				+ (double)drMarcelo["t4"]+(double)drMarcelo["t5"]+(double)drMarcelo["t6"]
				+ (double)drMarcelo["t7"]+(double)drMarcelo["t8"]+(double)drMarcelo["t9"]
				+ (double)drMarcelo["t10"]+(double)drMarcelo["t11"];	*/
		}

		private void txtTotalTarjetas_ValueChanged(object sender, System.EventArgs e)
		{
			//txtOtrasTarjetas.Value = + (double)txtTotalTarjetas.Value - (double)txtCuotaFacil.Value;
			//suma2();
			//096394652	
		}

		private void txtOtrasTarjetas_ValueChanged(object sender, System.EventArgs e)
		{
//			txtTotalTarjetas.Value = (double)txtCuotaFacil.Value + (double)txtOtrasTarjetas.Value;
			//suma2();
		}

		private void txtEfectivo_ValueChanged(object sender, System.EventArgs e)
		{
			//suma2();
		}

		private void txtCheque_ValueChanged(object sender, System.EventArgs e)
		{
			//suma2();
		}

		private void txtCuotaFacil_ValueChanged(object sender, System.EventArgs e)
		{
			//txtTotalTarjetas.Value = (double)txtCuotaFacil.Value + (double)txtOtrasTarjetas.Value;
			//suma2();
		}

		private void txtCredPersonal_ValueChanged(object sender, System.EventArgs e)
		{
			//suma2();
		}

		private void txtGastos_ValueChanged(object sender, System.EventArgs e)
		{
			//suma2();
		}

	
		private void txtHombre_Validated(object sender, System.EventArgs e)
		{
			suma();
		}

		private void txtMujer_Validated(object sender, System.EventArgs e)
		{
			suma();
		}

		private void txtNiño_Validated(object sender, System.EventArgs e)
		{
			suma();
		}

		private void txtNiña_Validated(object sender, System.EventArgs e)
		{
			suma();
		}

		private void txtCartera_Validated(object sender, System.EventArgs e)
		{
			suma();
		}

		private void txtCinturon_Validated(object sender, System.EventArgs e)
		{
			suma();
		}


		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			suma();
			suma2();
			barraDato1.GrabaRegistro();
			barraDato1.EditarRegistro();
			desactivar();
		}

		private void btDeshacer_Click(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();
			barraDato1.EditarRegistro();  
			desactivar();
		}

		private void desactivar()
		{
			txtOtrasTarjetas.Enabled=false;
			txtDeposito.Enabled=false;  
			txtValorActual.Enabled=false;
			txtValorAntes.Enabled=false;
  		txtUnidadActual.Enabled=false;
			txtUnidadesAntes.Enabled=false;  
			txtTotal.Enabled =false; 
		}

		private void txtGorra_Validated(object sender, System.EventArgs e)
		{
			suma();		
		}

		private void txtMochila_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtMaleta_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtCanguro_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtBilletera_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtPortazapatos_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtBalon_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtPromo5_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtPromo10_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtPromo15_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtPromo20_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtPromo25_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtPromo29_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtPromoOtro_Validated(object sender, System.EventArgs e)
		{
					suma();
		}

		private void txtEfectivo_Validated(object sender, System.EventArgs e)
		{
				suma2();
		}

		private void txtCheque_Validated(object sender, System.EventArgs e)
		{
				suma2();
		}

		private void txtCuotaFacil_Validated(object sender, System.EventArgs e)
		{
				drMarcelo["Tarjeta"]=(double)txtTotalTarjetas.Value-(double)drMarcelo["CoutaFacil"];
				suma2();
		}		
		private void txtTotalTarjetas_Validated(object sender, System.EventArgs e)
		{
				drMarcelo["Tarjeta"]=(double)txtTotalTarjetas.Value-(double)drMarcelo["CoutaFacil"];
				suma2();
		}

		private void txtOtrasTarjetas_Validated(object sender, System.EventArgs e)
		{
				suma2();	
		}

		private void txtCredPersonal_Validated(object sender, System.EventArgs e)
		{
				suma2();
		}

		private void txtGastos_Validated(object sender, System.EventArgs e)
		{
				suma2();
		}

		private void btReportes_Click(object sender, System.EventArgs e)
		{
				MarceloReportes   miMarceloReportes = new MarceloReportes();
				miMarceloReportes.Show();
		}

		private void txtHombre_Enter(object sender, System.EventArgs e)
		{
			txtHombre.SelectAll(); 
		}

		private void txtMujer_Enter(object sender, System.EventArgs e)
		{
			txtMujer.SelectAll();  
		}

		private void txtEfectivo_Enter(object sender, System.EventArgs e)
		{
			txtEfectivo.SelectAll();  
		}

		private void txtCheque_Enter(object sender, System.EventArgs e)
		{
			txtCheque.SelectAll();  
		}

		private void txtCuotaFacil_Enter(object sender, System.EventArgs e)
		{
			txtCuotaFacil.SelectAll();  
		}

		private void txtTotalTarjetas_Enter(object sender, System.EventArgs e)
		{
				txtTotalTarjetas.SelectAll();  
		}

		private void txtCredPersonal_Enter(object sender, System.EventArgs e)
		{	
			txtCredPersonal.SelectAll();  
		}

		private void txtGastos_Enter(object sender, System.EventArgs e)
		{
			txtGastos.SelectAll(); 
			//txtGastos.SelectAll();  
		}

		private void txtNiño_Enter(object sender, System.EventArgs e)
		{
			txtNiño.SelectAll(); 
		}

		private void txtNiña_Enter(object sender, System.EventArgs e)
		{
			txtNiña.SelectAll(); 
		}

		private void txtCartera_Enter(object sender, System.EventArgs e)
		{
			txtCartera.SelectAll();  
		}

		private void txtCinturon_Enter(object sender, System.EventArgs e)
		{
			txtCinturon.SelectAll();  
		}

		private void txtGorra_Enter(object sender, System.EventArgs e)
		{
			txtGorra.SelectAll();  
		}

		private void txtMochila_Enter(object sender, System.EventArgs e)
		{
			txtMochila.SelectAll();  
		}

		private void txtMaleta_Enter(object sender, System.EventArgs e)
		{
			txtMaleta.SelectAll();  
		}

		private void txtCanguro_Enter(object sender, System.EventArgs e)
		{
			txtCanguro.SelectAll();  
		}

		private void txtBilletera_Enter(object sender, System.EventArgs e)
		{
			txtBilletera.SelectAll();  
		}

		private void txtPortazapatos_Enter(object sender, System.EventArgs e)
		{
			txtPortazapatos.SelectAll();  
		}

		private void txtBalon_Enter(object sender, System.EventArgs e)
		{
			txtBalon.SelectAll();  
		}

		private void txtPromo5_Enter(object sender, System.EventArgs e)
		{
			txtPromo5.SelectAll();  
		}

		private void txtPromo10_Enter(object sender, System.EventArgs e)
		{
			txtPromo10.SelectAll();  
		}

		private void txtPromo15_Enter(object sender, System.EventArgs e)
		{	
			txtPromo15.SelectAll();  
		}

		private void txtPromo20_Enter(object sender, System.EventArgs e)
		{
			txtPromo20.SelectAll();  
		}

		private void txtPromo25_Enter(object sender, System.EventArgs e)
		{
			txtPromo25.SelectAll();  
		}

		private void txtPromo29_Enter(object sender, System.EventArgs e)
		{
			txtPromo29.SelectAll();  
		}

		private void txtPromoOtro_Enter(object sender, System.EventArgs e)
		{
			txtPromoOtro.SelectAll(); 
		}

		private void label80_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label37_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label88_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor64_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor8_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor35_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor33_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor33_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor32_Validated(object sender, System.EventArgs e)
		{
			suma();
		}

		private void ultraNumericEditor31_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor30_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor29_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor28_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor27_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor26_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor25_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor24_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor23_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor53_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor52_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor51_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor50_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor49_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor48_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor47_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor46_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor45_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void label98_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor1_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor1.SelectAll(); 
		}

		private void ultraNumericEditor22_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor22.SelectAll(); 
		}

		private void ultraNumericEditor33_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor33.SelectAll(); 
		}

		private void ultraNumericEditor2_Enter(object sender, System.EventArgs e)
		{

		}

		private void ultraNumericEditor21_Enter(object sender, System.EventArgs e)
		{

		}

		private void ultraNumericEditor32_Enter(object sender, System.EventArgs e)
		{

		}

		private void ultraNumericEditor3_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor3.SelectAll(); 
		}

		private void ultraNumericEditor20_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor20.SelectAll(); 
		}

		private void ultraNumericEditor31_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor31.SelectAll(); 
		}

		private void ultraNumericEditor42_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor42.SelectAll(); 
		}

		private void ultraNumericEditor53_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor53.SelectAll(); 
		}

		private void ultraNumericEditor4_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor4.SelectAll(); 
		}

		private void ultraNumericEditor19_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor19.SelectAll(); 
		}

		private void ultraNumericEditor30_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor30.SelectAll(); 
		}

		private void ultraNumericEditor41_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor41.SelectAll(); 
		}

		private void ultraNumericEditor52_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor52.SelectAll(); 
		}

		private void ultraNumericEditor5_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor5.SelectAll(); 
		}

		private void ultraNumericEditor18_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor18.SelectAll(); 
		}

		private void ultraNumericEditor29_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor29.SelectAll(); 
		}

		private void ultraNumericEditor40_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor40.SelectAll(); 
		}

		private void ultraNumericEditor51_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor51.SelectAll(); 
		}

		private void ultraNumericEditor6_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor6.SelectAll(); 
		}

		private void ultraNumericEditor17_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor17.SelectAll(); 
		}

		private void ultraNumericEditor28_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor28.SelectAll(); 
		}

		private void ultraNumericEditor39_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor39.SelectAll(); 
		}

		private void ultraNumericEditor50_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor50.SelectAll(); 
		}

		private void ultraNumericEditor7_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor7.SelectAll(); 
		}

		private void ultraNumericEditor16_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor16.SelectAll(); 
		}

		private void ultraNumericEditor27_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor27.SelectAll(); 
		}

		private void ultraNumericEditor38_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor38.SelectAll(); 
		}

		private void ultraNumericEditor49_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor49.SelectAll(); 
		}

		private void ultraNumericEditor8_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor8.SelectAll(); 
		}

		private void ultraNumericEditor15_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor15.SelectAll(); 
		}

		private void ultraNumericEditor26_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor26.SelectAll(); 
		}

		private void ultraNumericEditor37_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor37.SelectAll(); 
		}

		private void ultraNumericEditor48_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor48.SelectAll(); 
		}

		private void ultraNumericEditor9_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor9.SelectAll(); 
		}

		private void ultraNumericEditor14_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor14.SelectAll(); 
		}

		private void ultraNumericEditor25_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor25.SelectAll(); 
		}

		private void ultraNumericEditor36_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor36.SelectAll(); 
		}

		private void ultraNumericEditor47_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor47.SelectAll(); 
		}

		private void ultraNumericEditor10_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor10.SelectAll(); 
		}

		private void ultraNumericEditor13_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor13.SelectAll(); 
		}

		private void ultraNumericEditor24_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor24.SelectAll(); 
		}

		private void ultraNumericEditor35_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor35.SelectAll(); 
		}

		private void ultraNumericEditor46_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor46.SelectAll(); 
		}

		private void ultraNumericEditor11_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor11.SelectAll(); 
		}

		private void ultraNumericEditor12_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor12.SelectAll(); 
		}

		private void ultraNumericEditor23_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor23.SelectAll(); 
		}

		private void ultraNumericEditor34_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor34.SelectAll(); 
		}

		private void ultraNumericEditor45_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor45.SelectAll(); 
		}

		private void ultraNumericEditor45_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label96_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtRetencion_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void txtCredEmpresa_Validated(object sender, System.EventArgs e)
		{
		suma2();
		}

		private void txtRetencion_Enter(object sender, System.EventArgs e)
		{
			txtRetencion.SelectAll(); 
		}

		private void txtCredEmpresa_Enter(object sender, System.EventArgs e)
		{
			txtCredEmpresa.SelectAll(); 
		}

		private void txtPromo35_Enter(object sender, System.EventArgs e)
		{
			txtPromo35.SelectAll();  
		}

		private void txtPromo79_Enter(object sender, System.EventArgs e)
		{
			txtPromo79.SelectAll();  
		}

		private void txtPromo35_Validated(object sender, System.EventArgs e)
		{
			suma();
		}

		private void txtPromo79_Validated(object sender, System.EventArgs e)
		{
			suma();
		}

		private void txtIniciaF_Enter(object sender, System.EventArgs e)
		{
			txtIniciaF.SelectAll(); 
		}

		private void txtTerminaF_Enter(object sender, System.EventArgs e)
		{
			txtTerminaF.SelectAll();
		}

		private void txtIniciaNC_Enter(object sender, System.EventArgs e)
		{
			txtIniciaNC.SelectAll(); 
		}

		private void txtTerminaNC_Enter(object sender, System.EventArgs e)
		{
			txtTerminaNC.SelectAll(); 
		}

		private void MarceloGeneralVentas_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor42_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraCombo7_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraCombo10_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraNumericEditor41_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor39_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor38_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor51_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor48_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor46_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraCombo12_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraCombo14_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraCombo16_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraCombo17_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraCombo19_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraNumericEditor56_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor70_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor55_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor59_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor65_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor64_ValueChanged_1(object sender, System.EventArgs e)
		{
		
		}

		private void label71_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor21_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor43_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor44_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor55_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor57_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor59_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor61_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor63_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor65_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor67_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor69_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraNumericEditor21_Enter_1(object sender, System.EventArgs e)
		{
			ultraNumericEditor21.SelectAll(); 
		}

		private void ultraNumericEditor43_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor43.SelectAll(); 
		}

		private void ultraNumericEditor44_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor44.SelectAll(); 
		}

		private void ultraNumericEditor55_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor55.SelectAll(); 
		}

		private void ultraNumericEditor57_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor57.SelectAll(); 
		}

		private void ultraNumericEditor59_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor59.SelectAll(); 
		}

		private void ultraNumericEditor61_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor61.SelectAll(); 
		}

		private void ultraNumericEditor63_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor63.SelectAll(); 
		}

		private void ultraNumericEditor65_Enter(object sender, System.EventArgs e)
		{	
			ultraNumericEditor65.SelectAll(); 
		}

		private void ultraNumericEditor67_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor67.SelectAll(); 
		}

		private void ultraNumericEditor69_Enter(object sender, System.EventArgs e)
		{
			ultraNumericEditor69.SelectAll(); 
		}

		private void ultraCombo1_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo2_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo3_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo4_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo5_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo6_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo7_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo8_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo9_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo10_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo11_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo12_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo13_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo14_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo15_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo16_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo17_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo18_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo19_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void ultraCombo20_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void cet1_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void cet2_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void cet3_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void cet4_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void cet5_Validated(object sender, System.EventArgs e)
		{
			suma2();
		}

		private void cet1_Enter(object sender, System.EventArgs e)
		{
			cet1.SelectAll(); 
		}

		private void cet2_Enter(object sender, System.EventArgs e)
		{
			cet2.SelectAll(); 
		}

		private void cet3_Enter(object sender, System.EventArgs e)
		{
			cet3.SelectAll(); 
		}

		private void cet4_Enter(object sender, System.EventArgs e)
		{
			cet4.SelectAll(); 
		}

		private void cet5_Enter(object sender, System.EventArgs e)
		{
			cet5.SelectAll(); 
		}

		private void cefact1_Enter(object sender, System.EventArgs e)
		{
			cefact1.SelectAll(); 
		}

		private void cefact2_Enter(object sender, System.EventArgs e)
		{
			cefact2.SelectAll(); 	
		}

		private void cefact3_Enter(object sender, System.EventArgs e)
		{
			cefact3.SelectAll(); 
		}

		private void cefact4_Enter(object sender, System.EventArgs e)
		{
			cefact4.SelectAll(); 
		}

		private void cefact5_Enter(object sender, System.EventArgs e)
		{
			cefact5.SelectAll(); 
		}

		private void txtNombre1_Enter(object sender, System.EventArgs e)
		{
			txtNombre1.SelectAll(); 
		}

		private void txtNombre2_Enter(object sender, System.EventArgs e)
		{
			txtNombre2.SelectAll(); 
		}

		private void txtNombre3_Enter(object sender, System.EventArgs e)
		{
			txtNombre3.SelectAll(); 
		}

		private void txtNombre4_Enter(object sender, System.EventArgs e)
		{
			txtNombre4.SelectAll(); 
		}

		private void txtNombre5_Enter(object sender, System.EventArgs e)
		{
			txtNombre5.SelectAll(); 
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdsFormaPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsFormaPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}	

	}
}
