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
	/// Descripción breve de frmCuadriculaHabitaciones.
	/// </summary>
	public class frmCuadriculaHabitaciones : System.Windows.Forms.Form
	{
		#region detalle habitaciones
		
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.Button button4;
		private System.Windows.Forms.Button button5;
		private System.Windows.Forms.Button button6;
		private System.Windows.Forms.Button button7;
		private System.Windows.Forms.Button button8;
		private System.Windows.Forms.Button button9;
		private System.Windows.Forms.Button button10;
		private System.Windows.Forms.Button button11;
		private System.Windows.Forms.Button button12;
		private System.Windows.Forms.Button button13;
		private System.Windows.Forms.Button button14;
		private System.Windows.Forms.Button button15;
		private System.Windows.Forms.Button button16;
		private System.Windows.Forms.Button button17;
		private System.Windows.Forms.Button button18;
		private System.Windows.Forms.Button button19;
		private System.Windows.Forms.Button button20;
		private System.Windows.Forms.Button button21;
		private System.Windows.Forms.Button button22;
		private System.Windows.Forms.Button button23;
		private System.Windows.Forms.Button button24;
		private System.Windows.Forms.Button button25;
		private System.Windows.Forms.Button button26;
		private System.Windows.Forms.Button button27;
		private System.Windows.Forms.Button button28;
		private System.Windows.Forms.Button button29;
		private System.Windows.Forms.Button button30;
		private System.Windows.Forms.Button button31;
		private System.Windows.Forms.Button button32;
		private System.Windows.Forms.Button button33;
		private System.Windows.Forms.Button button34;
		private System.Windows.Forms.Button button35;
		private System.Windows.Forms.Button button36;
		private System.Windows.Forms.Button button37;
		private System.Windows.Forms.Button button38;
		private System.Windows.Forms.Button button39;
		private System.Windows.Forms.Button button40;
		private System.Windows.Forms.Button button41;
		private System.Windows.Forms.Button button42;
		private System.Windows.Forms.Button button43;
		private System.Windows.Forms.Button button44;
		private System.Windows.Forms.Button button45;
		private System.Windows.Forms.Button button46;
		private System.Windows.Forms.Button button47;
		private System.Windows.Forms.Button button48;
		private System.Windows.Forms.Label labH1;
		private System.Windows.Forms.Label labH2;
		private System.Windows.Forms.Label labH3;
		private System.Windows.Forms.Label labH6;
		private System.Windows.Forms.Label labH5;
		private System.Windows.Forms.Label labH4;
		private System.Windows.Forms.Label labH9;
		private System.Windows.Forms.Label labH8;
		private System.Windows.Forms.Label labH7;
		private System.Windows.Forms.Label labH12;
		private System.Windows.Forms.Label labH11;
		private System.Windows.Forms.Label labH10;
		private System.Windows.Forms.Label labH15;
		private System.Windows.Forms.Label labH14;
		private System.Windows.Forms.Label labH13;
		private System.Windows.Forms.Label labH18;
		private System.Windows.Forms.Label labH17;
		private System.Windows.Forms.Label labH16;
		private System.Windows.Forms.Label labH21;
		private System.Windows.Forms.Label labH20;
		private System.Windows.Forms.Label labH19;
		private System.Windows.Forms.Label labH24;
		private System.Windows.Forms.Label labH23;
		private System.Windows.Forms.Label labH22;
		private System.Windows.Forms.Label labH27;
		private System.Windows.Forms.Label labH26;
		private System.Windows.Forms.Label labH25;
		private System.Windows.Forms.Label labH30;
		private System.Windows.Forms.Label labH29;
		private System.Windows.Forms.Label labH28;
		private System.Windows.Forms.Label labH33;
		private System.Windows.Forms.Label labH32;
		private System.Windows.Forms.Label labH31;
		private System.Windows.Forms.Label labH36;
		private System.Windows.Forms.Label labH35;
		private System.Windows.Forms.Label labH34;
		private System.Windows.Forms.Label labH39;
		private System.Windows.Forms.Label labH38;
		private System.Windows.Forms.Label labH37;
		private System.Windows.Forms.Label labH42;
		private System.Windows.Forms.Label labH41;
		private System.Windows.Forms.Label labH40;
		private System.Windows.Forms.Label labH45;
		private System.Windows.Forms.Label labH44;
		private System.Windows.Forms.Label labH43;
		private System.Windows.Forms.Label labH48;
		private System.Windows.Forms.Label labH47;
		private System.Windows.Forms.Label labH46;
		private System.Windows.Forms.Label labH51;
		private System.Windows.Forms.Label labH50;
		private System.Windows.Forms.Label labH49;
		private System.Windows.Forms.Label labH54;
		private System.Windows.Forms.Label labH53;
		private System.Windows.Forms.Label labH52;
		private System.Windows.Forms.Label labH57;
		private System.Windows.Forms.Label labH56;
		private System.Windows.Forms.Label labH55;
		private System.Windows.Forms.Label labH60;
		private System.Windows.Forms.Label labH59;
		private System.Windows.Forms.Label labH58;
		private System.Windows.Forms.Label labH63;
		private System.Windows.Forms.Label labH62;
		private System.Windows.Forms.Label labH61;
		private System.Windows.Forms.Label labH66;
		private System.Windows.Forms.Label labH65;
		private System.Windows.Forms.Label labH64;
		private System.Windows.Forms.Label labH69;
		private System.Windows.Forms.Label labH68;
		private System.Windows.Forms.Label labH67;
		private System.Windows.Forms.Label labH72;
		private System.Windows.Forms.Label labH71;
		private System.Windows.Forms.Label labH70;
		private System.Windows.Forms.Label labH75;
		private System.Windows.Forms.Label labH74;
		private System.Windows.Forms.Label labH73;
		private System.Windows.Forms.Label labH78;
		private System.Windows.Forms.Label labH77;
		private System.Windows.Forms.Label labH76;
		private System.Windows.Forms.Label labH81;
		private System.Windows.Forms.Label labH80;
		private System.Windows.Forms.Label labH79;
		private System.Windows.Forms.Label labH84;
		private System.Windows.Forms.Label labH83;
		private System.Windows.Forms.Label labH82;
		private System.Windows.Forms.Label labH87;
		private System.Windows.Forms.Label labH86;
		private System.Windows.Forms.Label labH85;
		private System.Windows.Forms.Label labH90;
		private System.Windows.Forms.Label labH89;
		private System.Windows.Forms.Label labH88;
		private System.Windows.Forms.Label labH93;
		private System.Windows.Forms.Label labH92;
		private System.Windows.Forms.Label labH91;
		private System.Windows.Forms.Label labH96;
		private System.Windows.Forms.Label labH95;
		private System.Windows.Forms.Label labH94;
		private System.Windows.Forms.Label labH99;
		private System.Windows.Forms.Label labH98;
		private System.Windows.Forms.Label labH97;
		private System.Windows.Forms.Label labH102;
		private System.Windows.Forms.Label labH101;
		private System.Windows.Forms.Label labH100;
		private System.Windows.Forms.Label labH105;
		private System.Windows.Forms.Label labH104;
		private System.Windows.Forms.Label labH103;
		private System.Windows.Forms.Label labH108;
		private System.Windows.Forms.Label labH107;
		private System.Windows.Forms.Label labH106;
		private System.Windows.Forms.Label labH111;
		private System.Windows.Forms.Label labH110;
		private System.Windows.Forms.Label labH109;
		private System.Windows.Forms.Label labH114;
		private System.Windows.Forms.Label labH113;
		private System.Windows.Forms.Label labH112;
		private System.Windows.Forms.Label labH117;
		private System.Windows.Forms.Label labH116;
		private System.Windows.Forms.Label labH115;
		private System.Windows.Forms.Label labH120;
		private System.Windows.Forms.Label labH119;
		private System.Windows.Forms.Label labH118;
		private System.Windows.Forms.Label labH123;
		private System.Windows.Forms.Label labH122;
		private System.Windows.Forms.Label labH121;
		private System.Windows.Forms.Label labH126;
		private System.Windows.Forms.Label labH125;
		private System.Windows.Forms.Label labH124;
		private System.Windows.Forms.Label labH129;
		private System.Windows.Forms.Label labH128;
		private System.Windows.Forms.Label labH127;
		private System.Windows.Forms.Label labH132;
		private System.Windows.Forms.Label labH131;
		private System.Windows.Forms.Label labH130;
		private System.Windows.Forms.Label labH135;
		private System.Windows.Forms.Label labH134;
		private System.Windows.Forms.Label labH133;
		private System.Windows.Forms.Label labH138;
		private System.Windows.Forms.Label labH137;
		private System.Windows.Forms.Label labH136;
		private System.Windows.Forms.Label labH141;
		private System.Windows.Forms.Label labH140;
		private System.Windows.Forms.Label labH139;
		private System.Windows.Forms.Label labH144;
		private System.Windows.Forms.Label labH143;
		private System.Windows.Forms.Label labH142;
		#endregion detalle habitaciones
		private System.ComponentModel.IContainer components;


		string nombrdlBodega = "";
    int NumrHabtcn = 0;
		int idNroHab = 0;
		int iBodega = 0;
		int iRetTarifa = 0;
		private System.Windows.Forms.Timer timer1;
		private System.Windows.Forms.ContextMenu MnuEstado;
		int idBodega = 0;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.MenuItem mnuLimpieza;
		private System.Windows.Forms.MenuItem mnuLibre;
		private System.Windows.Forms.Timer timer2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label labelEn1;
		private System.Windows.Forms.Label labelEn2;
		private System.Windows.Forms.Label labelEn3;
		private System.Windows.Forms.Label labelEn4;
		private System.Windows.Forms.Label labelEn5;
		private System.Windows.Forms.Label labelEn6;
		private System.Windows.Forms.Label labelEn28;
		private System.Windows.Forms.Label labelEn27;
		private System.Windows.Forms.Label labelEn26;
		private System.Windows.Forms.Label labelEn25;
		private System.Windows.Forms.Label labelEn24;
		private System.Windows.Forms.Label labelEn23;
		private System.Windows.Forms.Label labelEn22;
		private System.Windows.Forms.Label labelEn21;
		private System.Windows.Forms.Label labelEn20;
		private System.Windows.Forms.Label labelEn19;
		private System.Windows.Forms.Label labelEn18;
		private System.Windows.Forms.Label labelEn17;
		private System.Windows.Forms.Label labelEn16;
		private System.Windows.Forms.Label labelEn15;
		private System.Windows.Forms.Label labelEn14;
		private System.Windows.Forms.Label labelEn13;
		private System.Windows.Forms.Label labelEn12;
		private System.Windows.Forms.Label labelEn11;
		private System.Windows.Forms.Label labelEn10;
		private System.Windows.Forms.Label labelEn9;
		private System.Windows.Forms.Label labelEn8;
		private System.Windows.Forms.Label labelEn7;
		private System.Windows.Forms.Label labelEn29;
		private System.Windows.Forms.Label labelEn30;
		private System.Windows.Forms.Label labelEn37;
		private System.Windows.Forms.Label labelEn36;
		private System.Windows.Forms.Label labelEn35;
		private System.Windows.Forms.Label labelEn34;
		private System.Windows.Forms.Label labelEn33;
		private System.Windows.Forms.Label labelEn32;
		private System.Windows.Forms.Label labelEn31;
		private System.Windows.Forms.Label labelEn38;
		private System.Windows.Forms.Label labelEn40;
		private System.Windows.Forms.Label labelEn39;
		private System.Windows.Forms.Label labelEn42;
		private System.Windows.Forms.Label labelEn41;
		private System.Windows.Forms.Label labelEn43;
		private System.Windows.Forms.Label labelEn48;
		private System.Windows.Forms.Label labelEn47;
		private System.Windows.Forms.Label labelEn46;
		private System.Windows.Forms.Label labelEn45;
		private System.Windows.Forms.Label labelEn44;
		private System.Windows.Forms.MenuItem menuIngresoCLI;
		private System.Windows.Forms.MenuItem menuEnvioCB;		
		int TiempoReser = 0;

		public frmCuadriculaHabitaciones()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCuadriculaHabitaciones));
			this.button1 = new System.Windows.Forms.Button();
			this.labH1 = new System.Windows.Forms.Label();
			this.labH2 = new System.Windows.Forms.Label();
			this.labH3 = new System.Windows.Forms.Label();
			this.labH6 = new System.Windows.Forms.Label();
			this.labH5 = new System.Windows.Forms.Label();
			this.labH4 = new System.Windows.Forms.Label();
			this.button2 = new System.Windows.Forms.Button();
			this.labH9 = new System.Windows.Forms.Label();
			this.labH8 = new System.Windows.Forms.Label();
			this.labH7 = new System.Windows.Forms.Label();
			this.button3 = new System.Windows.Forms.Button();
			this.labH12 = new System.Windows.Forms.Label();
			this.labH11 = new System.Windows.Forms.Label();
			this.labH10 = new System.Windows.Forms.Label();
			this.button4 = new System.Windows.Forms.Button();
			this.labH15 = new System.Windows.Forms.Label();
			this.labH14 = new System.Windows.Forms.Label();
			this.labH13 = new System.Windows.Forms.Label();
			this.button5 = new System.Windows.Forms.Button();
			this.labH18 = new System.Windows.Forms.Label();
			this.labH17 = new System.Windows.Forms.Label();
			this.labH16 = new System.Windows.Forms.Label();
			this.button6 = new System.Windows.Forms.Button();
			this.labH21 = new System.Windows.Forms.Label();
			this.labH20 = new System.Windows.Forms.Label();
			this.labH19 = new System.Windows.Forms.Label();
			this.button7 = new System.Windows.Forms.Button();
			this.labH24 = new System.Windows.Forms.Label();
			this.labH23 = new System.Windows.Forms.Label();
			this.labH22 = new System.Windows.Forms.Label();
			this.button8 = new System.Windows.Forms.Button();
			this.labH27 = new System.Windows.Forms.Label();
			this.labH26 = new System.Windows.Forms.Label();
			this.labH25 = new System.Windows.Forms.Label();
			this.button9 = new System.Windows.Forms.Button();
			this.labH30 = new System.Windows.Forms.Label();
			this.labH29 = new System.Windows.Forms.Label();
			this.labH28 = new System.Windows.Forms.Label();
			this.button10 = new System.Windows.Forms.Button();
			this.labH33 = new System.Windows.Forms.Label();
			this.labH32 = new System.Windows.Forms.Label();
			this.labH31 = new System.Windows.Forms.Label();
			this.button11 = new System.Windows.Forms.Button();
			this.labH36 = new System.Windows.Forms.Label();
			this.labH35 = new System.Windows.Forms.Label();
			this.labH34 = new System.Windows.Forms.Label();
			this.button12 = new System.Windows.Forms.Button();
			this.labH39 = new System.Windows.Forms.Label();
			this.labH38 = new System.Windows.Forms.Label();
			this.labH37 = new System.Windows.Forms.Label();
			this.button13 = new System.Windows.Forms.Button();
			this.labH42 = new System.Windows.Forms.Label();
			this.labH41 = new System.Windows.Forms.Label();
			this.labH40 = new System.Windows.Forms.Label();
			this.button14 = new System.Windows.Forms.Button();
			this.labH45 = new System.Windows.Forms.Label();
			this.labH44 = new System.Windows.Forms.Label();
			this.labH43 = new System.Windows.Forms.Label();
			this.button15 = new System.Windows.Forms.Button();
			this.labH48 = new System.Windows.Forms.Label();
			this.labH47 = new System.Windows.Forms.Label();
			this.labH46 = new System.Windows.Forms.Label();
			this.button16 = new System.Windows.Forms.Button();
			this.labH51 = new System.Windows.Forms.Label();
			this.labH50 = new System.Windows.Forms.Label();
			this.labH49 = new System.Windows.Forms.Label();
			this.button17 = new System.Windows.Forms.Button();
			this.labH54 = new System.Windows.Forms.Label();
			this.labH53 = new System.Windows.Forms.Label();
			this.labH52 = new System.Windows.Forms.Label();
			this.button18 = new System.Windows.Forms.Button();
			this.labH57 = new System.Windows.Forms.Label();
			this.labH56 = new System.Windows.Forms.Label();
			this.labH55 = new System.Windows.Forms.Label();
			this.button19 = new System.Windows.Forms.Button();
			this.labH60 = new System.Windows.Forms.Label();
			this.labH59 = new System.Windows.Forms.Label();
			this.labH58 = new System.Windows.Forms.Label();
			this.button20 = new System.Windows.Forms.Button();
			this.labH63 = new System.Windows.Forms.Label();
			this.labH62 = new System.Windows.Forms.Label();
			this.labH61 = new System.Windows.Forms.Label();
			this.button21 = new System.Windows.Forms.Button();
			this.labH66 = new System.Windows.Forms.Label();
			this.labH65 = new System.Windows.Forms.Label();
			this.labH64 = new System.Windows.Forms.Label();
			this.button22 = new System.Windows.Forms.Button();
			this.labH69 = new System.Windows.Forms.Label();
			this.labH68 = new System.Windows.Forms.Label();
			this.labH67 = new System.Windows.Forms.Label();
			this.button23 = new System.Windows.Forms.Button();
			this.labH72 = new System.Windows.Forms.Label();
			this.labH71 = new System.Windows.Forms.Label();
			this.labH70 = new System.Windows.Forms.Label();
			this.button24 = new System.Windows.Forms.Button();
			this.labH75 = new System.Windows.Forms.Label();
			this.labH74 = new System.Windows.Forms.Label();
			this.labH73 = new System.Windows.Forms.Label();
			this.button25 = new System.Windows.Forms.Button();
			this.labH78 = new System.Windows.Forms.Label();
			this.labH77 = new System.Windows.Forms.Label();
			this.labH76 = new System.Windows.Forms.Label();
			this.button26 = new System.Windows.Forms.Button();
			this.labH81 = new System.Windows.Forms.Label();
			this.labH80 = new System.Windows.Forms.Label();
			this.labH79 = new System.Windows.Forms.Label();
			this.button27 = new System.Windows.Forms.Button();
			this.labH84 = new System.Windows.Forms.Label();
			this.labH83 = new System.Windows.Forms.Label();
			this.labH82 = new System.Windows.Forms.Label();
			this.button28 = new System.Windows.Forms.Button();
			this.labH87 = new System.Windows.Forms.Label();
			this.labH86 = new System.Windows.Forms.Label();
			this.labH85 = new System.Windows.Forms.Label();
			this.button29 = new System.Windows.Forms.Button();
			this.labH90 = new System.Windows.Forms.Label();
			this.labH89 = new System.Windows.Forms.Label();
			this.labH88 = new System.Windows.Forms.Label();
			this.button30 = new System.Windows.Forms.Button();
			this.labH93 = new System.Windows.Forms.Label();
			this.labH92 = new System.Windows.Forms.Label();
			this.labH91 = new System.Windows.Forms.Label();
			this.button31 = new System.Windows.Forms.Button();
			this.labH96 = new System.Windows.Forms.Label();
			this.labH95 = new System.Windows.Forms.Label();
			this.labH94 = new System.Windows.Forms.Label();
			this.button32 = new System.Windows.Forms.Button();
			this.labH99 = new System.Windows.Forms.Label();
			this.labH98 = new System.Windows.Forms.Label();
			this.labH97 = new System.Windows.Forms.Label();
			this.button33 = new System.Windows.Forms.Button();
			this.labH102 = new System.Windows.Forms.Label();
			this.labH101 = new System.Windows.Forms.Label();
			this.labH100 = new System.Windows.Forms.Label();
			this.button34 = new System.Windows.Forms.Button();
			this.labH105 = new System.Windows.Forms.Label();
			this.labH104 = new System.Windows.Forms.Label();
			this.labH103 = new System.Windows.Forms.Label();
			this.button35 = new System.Windows.Forms.Button();
			this.labH108 = new System.Windows.Forms.Label();
			this.labH107 = new System.Windows.Forms.Label();
			this.labH106 = new System.Windows.Forms.Label();
			this.button36 = new System.Windows.Forms.Button();
			this.labH111 = new System.Windows.Forms.Label();
			this.labH110 = new System.Windows.Forms.Label();
			this.labH109 = new System.Windows.Forms.Label();
			this.button37 = new System.Windows.Forms.Button();
			this.labH114 = new System.Windows.Forms.Label();
			this.labH113 = new System.Windows.Forms.Label();
			this.labH112 = new System.Windows.Forms.Label();
			this.button38 = new System.Windows.Forms.Button();
			this.labH117 = new System.Windows.Forms.Label();
			this.labH116 = new System.Windows.Forms.Label();
			this.labH115 = new System.Windows.Forms.Label();
			this.button39 = new System.Windows.Forms.Button();
			this.labH120 = new System.Windows.Forms.Label();
			this.labH119 = new System.Windows.Forms.Label();
			this.labH118 = new System.Windows.Forms.Label();
			this.button40 = new System.Windows.Forms.Button();
			this.labH123 = new System.Windows.Forms.Label();
			this.labH122 = new System.Windows.Forms.Label();
			this.labH121 = new System.Windows.Forms.Label();
			this.button41 = new System.Windows.Forms.Button();
			this.labH126 = new System.Windows.Forms.Label();
			this.labH125 = new System.Windows.Forms.Label();
			this.labH124 = new System.Windows.Forms.Label();
			this.button42 = new System.Windows.Forms.Button();
			this.labH129 = new System.Windows.Forms.Label();
			this.labH128 = new System.Windows.Forms.Label();
			this.labH127 = new System.Windows.Forms.Label();
			this.button43 = new System.Windows.Forms.Button();
			this.labH132 = new System.Windows.Forms.Label();
			this.labH131 = new System.Windows.Forms.Label();
			this.labH130 = new System.Windows.Forms.Label();
			this.button44 = new System.Windows.Forms.Button();
			this.labH135 = new System.Windows.Forms.Label();
			this.labH134 = new System.Windows.Forms.Label();
			this.labH133 = new System.Windows.Forms.Label();
			this.button45 = new System.Windows.Forms.Button();
			this.labH138 = new System.Windows.Forms.Label();
			this.labH137 = new System.Windows.Forms.Label();
			this.labH136 = new System.Windows.Forms.Label();
			this.button46 = new System.Windows.Forms.Button();
			this.labH141 = new System.Windows.Forms.Label();
			this.labH140 = new System.Windows.Forms.Label();
			this.labH139 = new System.Windows.Forms.Label();
			this.button47 = new System.Windows.Forms.Button();
			this.labH144 = new System.Windows.Forms.Label();
			this.labH143 = new System.Windows.Forms.Label();
			this.labH142 = new System.Windows.Forms.Label();
			this.button48 = new System.Windows.Forms.Button();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.MnuEstado = new System.Windows.Forms.ContextMenu();
			this.mnuLimpieza = new System.Windows.Forms.MenuItem();
			this.mnuLibre = new System.Windows.Forms.MenuItem();
			this.menuIngresoCLI = new System.Windows.Forms.MenuItem();
			this.menuEnvioCB = new System.Windows.Forms.MenuItem();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.timer2 = new System.Windows.Forms.Timer(this.components);
			this.labelEn1 = new System.Windows.Forms.Label();
			this.labelEn2 = new System.Windows.Forms.Label();
			this.labelEn3 = new System.Windows.Forms.Label();
			this.labelEn4 = new System.Windows.Forms.Label();
			this.labelEn5 = new System.Windows.Forms.Label();
			this.labelEn6 = new System.Windows.Forms.Label();
			this.labelEn29 = new System.Windows.Forms.Label();
			this.labelEn28 = new System.Windows.Forms.Label();
			this.labelEn27 = new System.Windows.Forms.Label();
			this.labelEn26 = new System.Windows.Forms.Label();
			this.labelEn25 = new System.Windows.Forms.Label();
			this.labelEn24 = new System.Windows.Forms.Label();
			this.labelEn23 = new System.Windows.Forms.Label();
			this.labelEn22 = new System.Windows.Forms.Label();
			this.labelEn21 = new System.Windows.Forms.Label();
			this.labelEn20 = new System.Windows.Forms.Label();
			this.labelEn19 = new System.Windows.Forms.Label();
			this.labelEn18 = new System.Windows.Forms.Label();
			this.labelEn17 = new System.Windows.Forms.Label();
			this.labelEn16 = new System.Windows.Forms.Label();
			this.labelEn15 = new System.Windows.Forms.Label();
			this.labelEn14 = new System.Windows.Forms.Label();
			this.labelEn13 = new System.Windows.Forms.Label();
			this.labelEn12 = new System.Windows.Forms.Label();
			this.labelEn11 = new System.Windows.Forms.Label();
			this.labelEn10 = new System.Windows.Forms.Label();
			this.labelEn9 = new System.Windows.Forms.Label();
			this.labelEn8 = new System.Windows.Forms.Label();
			this.labelEn7 = new System.Windows.Forms.Label();
			this.labelEn43 = new System.Windows.Forms.Label();
			this.labelEn42 = new System.Windows.Forms.Label();
			this.labelEn41 = new System.Windows.Forms.Label();
			this.labelEn40 = new System.Windows.Forms.Label();
			this.labelEn39 = new System.Windows.Forms.Label();
			this.labelEn38 = new System.Windows.Forms.Label();
			this.labelEn37 = new System.Windows.Forms.Label();
			this.labelEn36 = new System.Windows.Forms.Label();
			this.labelEn35 = new System.Windows.Forms.Label();
			this.labelEn34 = new System.Windows.Forms.Label();
			this.labelEn33 = new System.Windows.Forms.Label();
			this.labelEn32 = new System.Windows.Forms.Label();
			this.labelEn31 = new System.Windows.Forms.Label();
			this.labelEn30 = new System.Windows.Forms.Label();
			this.labelEn48 = new System.Windows.Forms.Label();
			this.labelEn47 = new System.Windows.Forms.Label();
			this.labelEn46 = new System.Windows.Forms.Label();
			this.labelEn45 = new System.Windows.Forms.Label();
			this.labelEn44 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// button1
			// 
			this.button1.BackColor = System.Drawing.Color.Green;
			this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button1.ForeColor = System.Drawing.Color.White;
			this.button1.Location = new System.Drawing.Point(8, 8);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(152, 88);
			this.button1.TabIndex = 1;
			this.button1.Text = "Habitacion # 1";
			this.button1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			this.button1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button1_MouseDown);
			// 
			// labH1
			// 
			this.labH1.BackColor = System.Drawing.Color.Green;
			this.labH1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH1.ForeColor = System.Drawing.Color.White;
			this.labH1.Location = new System.Drawing.Point(32, 48);
			this.labH1.Name = "labH1";
			this.labH1.Size = new System.Drawing.Size(48, 16);
			this.labH1.TabIndex = 1;
			this.labH1.Text = "Hora 1";
			this.labH1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labH1.Click += new System.EventHandler(this.labH1_Click);
			// 
			// labH2
			// 
			this.labH2.BackColor = System.Drawing.Color.Green;
			this.labH2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH2.ForeColor = System.Drawing.Color.White;
			this.labH2.Location = new System.Drawing.Point(88, 48);
			this.labH2.Name = "labH2";
			this.labH2.Size = new System.Drawing.Size(48, 16);
			this.labH2.TabIndex = 2;
			this.labH2.Text = "Hora 2";
			this.labH2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labH2.Click += new System.EventHandler(this.labH2_Click);
			// 
			// labH3
			// 
			this.labH3.BackColor = System.Drawing.Color.Green;
			this.labH3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH3.ForeColor = System.Drawing.Color.White;
			this.labH3.Location = new System.Drawing.Point(24, 72);
			this.labH3.Name = "labH3";
			this.labH3.Size = new System.Drawing.Size(120, 16);
			this.labH3.TabIndex = 3;
			this.labH3.Text = "Estado";
			this.labH3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labH3.Click += new System.EventHandler(this.labH3_Click);
			// 
			// labH6
			// 
			this.labH6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH6.Location = new System.Drawing.Point(176, 72);
			this.labH6.Name = "labH6";
			this.labH6.Size = new System.Drawing.Size(120, 16);
			this.labH6.TabIndex = 6;
			this.labH6.Text = "Estado";
			this.labH6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH5
			// 
			this.labH5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH5.Location = new System.Drawing.Point(240, 48);
			this.labH5.Name = "labH5";
			this.labH5.Size = new System.Drawing.Size(48, 16);
			this.labH5.TabIndex = 5;
			this.labH5.Text = "Hora 2";
			this.labH5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH4
			// 
			this.labH4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH4.Location = new System.Drawing.Point(184, 48);
			this.labH4.Name = "labH4";
			this.labH4.Size = new System.Drawing.Size(48, 16);
			this.labH4.TabIndex = 4;
			this.labH4.Text = "Hora 1";
			this.labH4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button2
			// 
			this.button2.BackColor = System.Drawing.Color.Green;
			this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button2.Location = new System.Drawing.Point(160, 8);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(152, 88);
			this.button2.TabIndex = 2;
			this.button2.Text = "Habitacion # 1";
			this.button2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			this.button2.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button2_MouseDown);
			// 
			// labH9
			// 
			this.labH9.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH9.Location = new System.Drawing.Point(328, 72);
			this.labH9.Name = "labH9";
			this.labH9.Size = new System.Drawing.Size(120, 16);
			this.labH9.TabIndex = 9;
			this.labH9.Text = "Estado";
			this.labH9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH8
			// 
			this.labH8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH8.Location = new System.Drawing.Point(392, 48);
			this.labH8.Name = "labH8";
			this.labH8.Size = new System.Drawing.Size(48, 16);
			this.labH8.TabIndex = 8;
			this.labH8.Text = "Hora 2";
			this.labH8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH7
			// 
			this.labH7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH7.Location = new System.Drawing.Point(336, 48);
			this.labH7.Name = "labH7";
			this.labH7.Size = new System.Drawing.Size(48, 16);
			this.labH7.TabIndex = 7;
			this.labH7.Text = "Hora 1";
			this.labH7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button3
			// 
			this.button3.BackColor = System.Drawing.Color.SaddleBrown;
			this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button3.Location = new System.Drawing.Point(312, 8);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(152, 88);
			this.button3.TabIndex = 3;
			this.button3.Text = "Habitacion # 1";
			this.button3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			this.button3.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button3_MouseDown);
			// 
			// labH12
			// 
			this.labH12.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH12.Location = new System.Drawing.Point(480, 72);
			this.labH12.Name = "labH12";
			this.labH12.Size = new System.Drawing.Size(120, 16);
			this.labH12.TabIndex = 12;
			this.labH12.Text = "Estado";
			this.labH12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH11
			// 
			this.labH11.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH11.Location = new System.Drawing.Point(544, 48);
			this.labH11.Name = "labH11";
			this.labH11.Size = new System.Drawing.Size(48, 16);
			this.labH11.TabIndex = 11;
			this.labH11.Text = "Hora 2";
			this.labH11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH10
			// 
			this.labH10.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH10.Location = new System.Drawing.Point(488, 48);
			this.labH10.Name = "labH10";
			this.labH10.Size = new System.Drawing.Size(48, 16);
			this.labH10.TabIndex = 10;
			this.labH10.Text = "Hora 1";
			this.labH10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button4
			// 
			this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button4.Location = new System.Drawing.Point(464, 8);
			this.button4.Name = "button4";
			this.button4.Size = new System.Drawing.Size(152, 88);
			this.button4.TabIndex = 4;
			this.button4.Text = "Habitacion # 1";
			this.button4.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button4.Click += new System.EventHandler(this.button4_Click);
			this.button4.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button4_MouseDown);
			// 
			// labH15
			// 
			this.labH15.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH15.Location = new System.Drawing.Point(632, 72);
			this.labH15.Name = "labH15";
			this.labH15.Size = new System.Drawing.Size(120, 16);
			this.labH15.TabIndex = 15;
			this.labH15.Text = "Estado";
			this.labH15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH14
			// 
			this.labH14.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH14.Location = new System.Drawing.Point(696, 48);
			this.labH14.Name = "labH14";
			this.labH14.Size = new System.Drawing.Size(48, 16);
			this.labH14.TabIndex = 14;
			this.labH14.Text = "Hora 2";
			this.labH14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH13
			// 
			this.labH13.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH13.Location = new System.Drawing.Point(640, 48);
			this.labH13.Name = "labH13";
			this.labH13.Size = new System.Drawing.Size(48, 16);
			this.labH13.TabIndex = 13;
			this.labH13.Text = "Hora 1";
			this.labH13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button5
			// 
			this.button5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button5.Location = new System.Drawing.Point(616, 8);
			this.button5.Name = "button5";
			this.button5.Size = new System.Drawing.Size(152, 88);
			this.button5.TabIndex = 16;
			this.button5.Text = "Habitacion # 1";
			this.button5.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button5.Click += new System.EventHandler(this.button5_Click);
			this.button5.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button5_MouseDown);
			// 
			// labH18
			// 
			this.labH18.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH18.Location = new System.Drawing.Point(784, 72);
			this.labH18.Name = "labH18";
			this.labH18.Size = new System.Drawing.Size(120, 16);
			this.labH18.TabIndex = 18;
			this.labH18.Text = "Estado";
			this.labH18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH17
			// 
			this.labH17.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH17.Location = new System.Drawing.Point(848, 48);
			this.labH17.Name = "labH17";
			this.labH17.Size = new System.Drawing.Size(48, 16);
			this.labH17.TabIndex = 17;
			this.labH17.Text = "Hora 2";
			this.labH17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH16
			// 
			this.labH16.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH16.Location = new System.Drawing.Point(792, 48);
			this.labH16.Name = "labH16";
			this.labH16.Size = new System.Drawing.Size(48, 16);
			this.labH16.TabIndex = 16;
			this.labH16.Text = "Hora 1";
			this.labH16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button6
			// 
			this.button6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button6.Location = new System.Drawing.Point(768, 8);
			this.button6.Name = "button6";
			this.button6.Size = new System.Drawing.Size(152, 88);
			this.button6.TabIndex = 20;
			this.button6.Text = "Habitacion # 1";
			this.button6.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button6.Click += new System.EventHandler(this.button6_Click);
			this.button6.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button6_MouseDown);
			// 
			// labH21
			// 
			this.labH21.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH21.Location = new System.Drawing.Point(936, 72);
			this.labH21.Name = "labH21";
			this.labH21.Size = new System.Drawing.Size(120, 16);
			this.labH21.TabIndex = 21;
			this.labH21.Text = "Estado";
			this.labH21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH20
			// 
			this.labH20.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH20.Location = new System.Drawing.Point(1000, 48);
			this.labH20.Name = "labH20";
			this.labH20.Size = new System.Drawing.Size(48, 16);
			this.labH20.TabIndex = 20;
			this.labH20.Text = "Hora 2";
			this.labH20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH19
			// 
			this.labH19.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH19.Location = new System.Drawing.Point(944, 48);
			this.labH19.Name = "labH19";
			this.labH19.Size = new System.Drawing.Size(48, 16);
			this.labH19.TabIndex = 19;
			this.labH19.Text = "Hora 1";
			this.labH19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button7
			// 
			this.button7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button7.Location = new System.Drawing.Point(920, 8);
			this.button7.Name = "button7";
			this.button7.Size = new System.Drawing.Size(152, 88);
			this.button7.TabIndex = 24;
			this.button7.Text = "Habitacion # 1";
			this.button7.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button7.Click += new System.EventHandler(this.button7_Click);
			this.button7.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button7_MouseDown);
			// 
			// labH24
			// 
			this.labH24.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH24.Location = new System.Drawing.Point(1088, 72);
			this.labH24.Name = "labH24";
			this.labH24.Size = new System.Drawing.Size(120, 16);
			this.labH24.TabIndex = 24;
			this.labH24.Text = "Estado";
			this.labH24.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH23
			// 
			this.labH23.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH23.Location = new System.Drawing.Point(1152, 48);
			this.labH23.Name = "labH23";
			this.labH23.Size = new System.Drawing.Size(48, 16);
			this.labH23.TabIndex = 23;
			this.labH23.Text = "Hora 2";
			this.labH23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH22
			// 
			this.labH22.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH22.Location = new System.Drawing.Point(1096, 48);
			this.labH22.Name = "labH22";
			this.labH22.Size = new System.Drawing.Size(48, 16);
			this.labH22.TabIndex = 22;
			this.labH22.Text = "Hora 1";
			this.labH22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button8
			// 
			this.button8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button8.Location = new System.Drawing.Point(1072, 8);
			this.button8.Name = "button8";
			this.button8.Size = new System.Drawing.Size(152, 88);
			this.button8.TabIndex = 28;
			this.button8.Text = "Habitacion # 1";
			this.button8.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button8.Click += new System.EventHandler(this.button8_Click);
			this.button8.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button8_MouseDown);
			// 
			// labH27
			// 
			this.labH27.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH27.Location = new System.Drawing.Point(24, 160);
			this.labH27.Name = "labH27";
			this.labH27.Size = new System.Drawing.Size(120, 16);
			this.labH27.TabIndex = 27;
			this.labH27.Text = "Estado";
			this.labH27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH26
			// 
			this.labH26.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH26.Location = new System.Drawing.Point(88, 136);
			this.labH26.Name = "labH26";
			this.labH26.Size = new System.Drawing.Size(48, 16);
			this.labH26.TabIndex = 26;
			this.labH26.Text = "Hora 2";
			this.labH26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH25
			// 
			this.labH25.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH25.Location = new System.Drawing.Point(32, 136);
			this.labH25.Name = "labH25";
			this.labH25.Size = new System.Drawing.Size(48, 16);
			this.labH25.TabIndex = 25;
			this.labH25.Text = "Hora 1";
			this.labH25.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button9
			// 
			this.button9.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button9.Location = new System.Drawing.Point(8, 96);
			this.button9.Name = "button9";
			this.button9.Size = new System.Drawing.Size(152, 88);
			this.button9.TabIndex = 32;
			this.button9.Text = "Habitacion # 1";
			this.button9.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button9.Click += new System.EventHandler(this.button9_Click);
			this.button9.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button9_MouseDown);
			// 
			// labH30
			// 
			this.labH30.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH30.Location = new System.Drawing.Point(176, 160);
			this.labH30.Name = "labH30";
			this.labH30.Size = new System.Drawing.Size(120, 16);
			this.labH30.TabIndex = 30;
			this.labH30.Text = "Estado";
			this.labH30.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH29
			// 
			this.labH29.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH29.Location = new System.Drawing.Point(240, 136);
			this.labH29.Name = "labH29";
			this.labH29.Size = new System.Drawing.Size(48, 16);
			this.labH29.TabIndex = 29;
			this.labH29.Text = "Hora 2";
			this.labH29.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH28
			// 
			this.labH28.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH28.Location = new System.Drawing.Point(184, 136);
			this.labH28.Name = "labH28";
			this.labH28.Size = new System.Drawing.Size(48, 16);
			this.labH28.TabIndex = 28;
			this.labH28.Text = "Hora 1";
			this.labH28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button10
			// 
			this.button10.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button10.Location = new System.Drawing.Point(160, 96);
			this.button10.Name = "button10";
			this.button10.Size = new System.Drawing.Size(152, 88);
			this.button10.TabIndex = 36;
			this.button10.Text = "Habitacion # 1";
			this.button10.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button10.Click += new System.EventHandler(this.button10_Click);
			this.button10.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button10_MouseDown);
			// 
			// labH33
			// 
			this.labH33.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH33.Location = new System.Drawing.Point(328, 160);
			this.labH33.Name = "labH33";
			this.labH33.Size = new System.Drawing.Size(120, 16);
			this.labH33.TabIndex = 33;
			this.labH33.Text = "Estado";
			this.labH33.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH32
			// 
			this.labH32.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH32.Location = new System.Drawing.Point(392, 136);
			this.labH32.Name = "labH32";
			this.labH32.Size = new System.Drawing.Size(48, 16);
			this.labH32.TabIndex = 32;
			this.labH32.Text = "Hora 2";
			this.labH32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH31
			// 
			this.labH31.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH31.Location = new System.Drawing.Point(336, 136);
			this.labH31.Name = "labH31";
			this.labH31.Size = new System.Drawing.Size(48, 16);
			this.labH31.TabIndex = 31;
			this.labH31.Text = "Hora 1";
			this.labH31.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button11
			// 
			this.button11.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button11.Location = new System.Drawing.Point(312, 96);
			this.button11.Name = "button11";
			this.button11.Size = new System.Drawing.Size(152, 88);
			this.button11.TabIndex = 40;
			this.button11.Text = "Habitacion # 1";
			this.button11.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button11.Click += new System.EventHandler(this.button11_Click);
			this.button11.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button11_MouseDown);
			// 
			// labH36
			// 
			this.labH36.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH36.Location = new System.Drawing.Point(480, 160);
			this.labH36.Name = "labH36";
			this.labH36.Size = new System.Drawing.Size(120, 16);
			this.labH36.TabIndex = 36;
			this.labH36.Text = "Estado";
			this.labH36.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH35
			// 
			this.labH35.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH35.Location = new System.Drawing.Point(544, 136);
			this.labH35.Name = "labH35";
			this.labH35.Size = new System.Drawing.Size(48, 16);
			this.labH35.TabIndex = 35;
			this.labH35.Text = "Hora 2";
			this.labH35.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH34
			// 
			this.labH34.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH34.Location = new System.Drawing.Point(488, 136);
			this.labH34.Name = "labH34";
			this.labH34.Size = new System.Drawing.Size(48, 16);
			this.labH34.TabIndex = 34;
			this.labH34.Text = "Hora 1";
			this.labH34.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button12
			// 
			this.button12.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button12.Location = new System.Drawing.Point(464, 96);
			this.button12.Name = "button12";
			this.button12.Size = new System.Drawing.Size(152, 88);
			this.button12.TabIndex = 44;
			this.button12.Text = "Habitacion # 1";
			this.button12.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button12.Click += new System.EventHandler(this.button12_Click);
			this.button12.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button12_MouseDown);
			// 
			// labH39
			// 
			this.labH39.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH39.Location = new System.Drawing.Point(632, 160);
			this.labH39.Name = "labH39";
			this.labH39.Size = new System.Drawing.Size(120, 16);
			this.labH39.TabIndex = 39;
			this.labH39.Text = "Estado";
			this.labH39.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH38
			// 
			this.labH38.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH38.Location = new System.Drawing.Point(696, 136);
			this.labH38.Name = "labH38";
			this.labH38.Size = new System.Drawing.Size(48, 16);
			this.labH38.TabIndex = 38;
			this.labH38.Text = "Hora 2";
			this.labH38.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH37
			// 
			this.labH37.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH37.Location = new System.Drawing.Point(640, 136);
			this.labH37.Name = "labH37";
			this.labH37.Size = new System.Drawing.Size(48, 16);
			this.labH37.TabIndex = 37;
			this.labH37.Text = "Hora 1";
			this.labH37.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button13
			// 
			this.button13.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button13.Location = new System.Drawing.Point(616, 96);
			this.button13.Name = "button13";
			this.button13.Size = new System.Drawing.Size(152, 88);
			this.button13.TabIndex = 48;
			this.button13.Text = "Habitacion # 1";
			this.button13.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button13.Click += new System.EventHandler(this.button13_Click);
			this.button13.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button13_MouseDown);
			// 
			// labH42
			// 
			this.labH42.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH42.Location = new System.Drawing.Point(784, 160);
			this.labH42.Name = "labH42";
			this.labH42.Size = new System.Drawing.Size(120, 16);
			this.labH42.TabIndex = 42;
			this.labH42.Text = "Estado";
			this.labH42.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH41
			// 
			this.labH41.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH41.Location = new System.Drawing.Point(848, 136);
			this.labH41.Name = "labH41";
			this.labH41.Size = new System.Drawing.Size(48, 16);
			this.labH41.TabIndex = 41;
			this.labH41.Text = "Hora 2";
			this.labH41.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH40
			// 
			this.labH40.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH40.Location = new System.Drawing.Point(792, 136);
			this.labH40.Name = "labH40";
			this.labH40.Size = new System.Drawing.Size(48, 16);
			this.labH40.TabIndex = 40;
			this.labH40.Text = "Hora 1";
			this.labH40.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button14
			// 
			this.button14.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button14.Location = new System.Drawing.Point(768, 96);
			this.button14.Name = "button14";
			this.button14.Size = new System.Drawing.Size(152, 88);
			this.button14.TabIndex = 52;
			this.button14.Text = "Habitacion # 1";
			this.button14.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button14.Click += new System.EventHandler(this.button14_Click);
			this.button14.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button14_MouseDown);
			// 
			// labH45
			// 
			this.labH45.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH45.Location = new System.Drawing.Point(936, 160);
			this.labH45.Name = "labH45";
			this.labH45.Size = new System.Drawing.Size(120, 16);
			this.labH45.TabIndex = 45;
			this.labH45.Text = "Estado";
			this.labH45.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH44
			// 
			this.labH44.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH44.Location = new System.Drawing.Point(1000, 136);
			this.labH44.Name = "labH44";
			this.labH44.Size = new System.Drawing.Size(48, 16);
			this.labH44.TabIndex = 44;
			this.labH44.Text = "Hora 2";
			this.labH44.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH43
			// 
			this.labH43.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH43.Location = new System.Drawing.Point(944, 136);
			this.labH43.Name = "labH43";
			this.labH43.Size = new System.Drawing.Size(48, 16);
			this.labH43.TabIndex = 43;
			this.labH43.Text = "Hora 1";
			this.labH43.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button15
			// 
			this.button15.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button15.Location = new System.Drawing.Point(920, 96);
			this.button15.Name = "button15";
			this.button15.Size = new System.Drawing.Size(152, 88);
			this.button15.TabIndex = 56;
			this.button15.Text = "Habitacion # 1";
			this.button15.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button15.Click += new System.EventHandler(this.button15_Click);
			this.button15.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button15_MouseDown);
			// 
			// labH48
			// 
			this.labH48.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH48.Location = new System.Drawing.Point(1088, 160);
			this.labH48.Name = "labH48";
			this.labH48.Size = new System.Drawing.Size(120, 16);
			this.labH48.TabIndex = 48;
			this.labH48.Text = "Estado";
			this.labH48.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH47
			// 
			this.labH47.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH47.Location = new System.Drawing.Point(1152, 136);
			this.labH47.Name = "labH47";
			this.labH47.Size = new System.Drawing.Size(48, 16);
			this.labH47.TabIndex = 47;
			this.labH47.Text = "Hora 2";
			this.labH47.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH46
			// 
			this.labH46.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH46.Location = new System.Drawing.Point(1096, 136);
			this.labH46.Name = "labH46";
			this.labH46.Size = new System.Drawing.Size(48, 16);
			this.labH46.TabIndex = 46;
			this.labH46.Text = "Hora 1";
			this.labH46.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button16
			// 
			this.button16.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button16.Location = new System.Drawing.Point(1072, 96);
			this.button16.Name = "button16";
			this.button16.Size = new System.Drawing.Size(152, 88);
			this.button16.TabIndex = 60;
			this.button16.Text = "Habitacion # 1";
			this.button16.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button16.Click += new System.EventHandler(this.button16_Click);
			this.button16.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button16_MouseDown);
			// 
			// labH51
			// 
			this.labH51.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH51.Location = new System.Drawing.Point(24, 248);
			this.labH51.Name = "labH51";
			this.labH51.Size = new System.Drawing.Size(120, 16);
			this.labH51.TabIndex = 51;
			this.labH51.Text = "Estado";
			this.labH51.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH50
			// 
			this.labH50.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH50.Location = new System.Drawing.Point(88, 224);
			this.labH50.Name = "labH50";
			this.labH50.Size = new System.Drawing.Size(48, 16);
			this.labH50.TabIndex = 50;
			this.labH50.Text = "Hora 2";
			this.labH50.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH49
			// 
			this.labH49.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH49.Location = new System.Drawing.Point(32, 224);
			this.labH49.Name = "labH49";
			this.labH49.Size = new System.Drawing.Size(48, 16);
			this.labH49.TabIndex = 49;
			this.labH49.Text = "Hora 1";
			this.labH49.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button17
			// 
			this.button17.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button17.Location = new System.Drawing.Point(8, 184);
			this.button17.Name = "button17";
			this.button17.Size = new System.Drawing.Size(152, 88);
			this.button17.TabIndex = 64;
			this.button17.Text = "Habitacion # 1";
			this.button17.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button17.Click += new System.EventHandler(this.button17_Click);
			this.button17.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button17_MouseDown);
			// 
			// labH54
			// 
			this.labH54.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH54.Location = new System.Drawing.Point(176, 248);
			this.labH54.Name = "labH54";
			this.labH54.Size = new System.Drawing.Size(120, 16);
			this.labH54.TabIndex = 54;
			this.labH54.Text = "Estado";
			this.labH54.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH53
			// 
			this.labH53.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH53.Location = new System.Drawing.Point(240, 224);
			this.labH53.Name = "labH53";
			this.labH53.Size = new System.Drawing.Size(48, 16);
			this.labH53.TabIndex = 53;
			this.labH53.Text = "Hora 2";
			this.labH53.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH52
			// 
			this.labH52.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH52.Location = new System.Drawing.Point(184, 224);
			this.labH52.Name = "labH52";
			this.labH52.Size = new System.Drawing.Size(48, 16);
			this.labH52.TabIndex = 52;
			this.labH52.Text = "Hora 1";
			this.labH52.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button18
			// 
			this.button18.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button18.Location = new System.Drawing.Point(160, 184);
			this.button18.Name = "button18";
			this.button18.Size = new System.Drawing.Size(152, 88);
			this.button18.TabIndex = 68;
			this.button18.Text = "Habitacion # 1";
			this.button18.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button18.Click += new System.EventHandler(this.button18_Click);
			this.button18.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button18_MouseDown);
			// 
			// labH57
			// 
			this.labH57.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH57.Location = new System.Drawing.Point(328, 248);
			this.labH57.Name = "labH57";
			this.labH57.Size = new System.Drawing.Size(120, 16);
			this.labH57.TabIndex = 57;
			this.labH57.Text = "Estado";
			this.labH57.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH56
			// 
			this.labH56.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH56.Location = new System.Drawing.Point(392, 224);
			this.labH56.Name = "labH56";
			this.labH56.Size = new System.Drawing.Size(48, 16);
			this.labH56.TabIndex = 56;
			this.labH56.Text = "Hora 2";
			this.labH56.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH55
			// 
			this.labH55.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH55.Location = new System.Drawing.Point(336, 224);
			this.labH55.Name = "labH55";
			this.labH55.Size = new System.Drawing.Size(48, 16);
			this.labH55.TabIndex = 55;
			this.labH55.Text = "Hora 1";
			this.labH55.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button19
			// 
			this.button19.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button19.Location = new System.Drawing.Point(312, 184);
			this.button19.Name = "button19";
			this.button19.Size = new System.Drawing.Size(152, 88);
			this.button19.TabIndex = 72;
			this.button19.Text = "Habitacion # 1";
			this.button19.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button19.Click += new System.EventHandler(this.button19_Click);
			this.button19.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button19_MouseDown);
			// 
			// labH60
			// 
			this.labH60.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH60.Location = new System.Drawing.Point(480, 248);
			this.labH60.Name = "labH60";
			this.labH60.Size = new System.Drawing.Size(120, 16);
			this.labH60.TabIndex = 60;
			this.labH60.Text = "Estado";
			this.labH60.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH59
			// 
			this.labH59.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH59.Location = new System.Drawing.Point(544, 224);
			this.labH59.Name = "labH59";
			this.labH59.Size = new System.Drawing.Size(48, 16);
			this.labH59.TabIndex = 59;
			this.labH59.Text = "Hora 2";
			this.labH59.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH58
			// 
			this.labH58.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH58.Location = new System.Drawing.Point(488, 224);
			this.labH58.Name = "labH58";
			this.labH58.Size = new System.Drawing.Size(48, 16);
			this.labH58.TabIndex = 58;
			this.labH58.Text = "Hora 1";
			this.labH58.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button20
			// 
			this.button20.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button20.Location = new System.Drawing.Point(464, 184);
			this.button20.Name = "button20";
			this.button20.Size = new System.Drawing.Size(152, 88);
			this.button20.TabIndex = 76;
			this.button20.Text = "Habitacion # 1";
			this.button20.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button20.Click += new System.EventHandler(this.button20_Click);
			this.button20.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button20_MouseDown);
			// 
			// labH63
			// 
			this.labH63.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH63.Location = new System.Drawing.Point(632, 248);
			this.labH63.Name = "labH63";
			this.labH63.Size = new System.Drawing.Size(120, 16);
			this.labH63.TabIndex = 63;
			this.labH63.Text = "Estado";
			this.labH63.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH62
			// 
			this.labH62.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH62.Location = new System.Drawing.Point(696, 224);
			this.labH62.Name = "labH62";
			this.labH62.Size = new System.Drawing.Size(48, 16);
			this.labH62.TabIndex = 62;
			this.labH62.Text = "Hora 2";
			this.labH62.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH61
			// 
			this.labH61.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH61.Location = new System.Drawing.Point(640, 224);
			this.labH61.Name = "labH61";
			this.labH61.Size = new System.Drawing.Size(48, 16);
			this.labH61.TabIndex = 61;
			this.labH61.Text = "Hora 1";
			this.labH61.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button21
			// 
			this.button21.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button21.Location = new System.Drawing.Point(616, 184);
			this.button21.Name = "button21";
			this.button21.Size = new System.Drawing.Size(152, 88);
			this.button21.TabIndex = 80;
			this.button21.Text = "Habitacion # 1";
			this.button21.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button21.Click += new System.EventHandler(this.button21_Click);
			this.button21.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button21_MouseDown);
			// 
			// labH66
			// 
			this.labH66.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH66.Location = new System.Drawing.Point(784, 248);
			this.labH66.Name = "labH66";
			this.labH66.Size = new System.Drawing.Size(120, 16);
			this.labH66.TabIndex = 66;
			this.labH66.Text = "Estado";
			this.labH66.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH65
			// 
			this.labH65.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH65.Location = new System.Drawing.Point(848, 224);
			this.labH65.Name = "labH65";
			this.labH65.Size = new System.Drawing.Size(48, 16);
			this.labH65.TabIndex = 65;
			this.labH65.Text = "Hora 2";
			this.labH65.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH64
			// 
			this.labH64.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH64.Location = new System.Drawing.Point(792, 224);
			this.labH64.Name = "labH64";
			this.labH64.Size = new System.Drawing.Size(48, 16);
			this.labH64.TabIndex = 64;
			this.labH64.Text = "Hora 1";
			this.labH64.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button22
			// 
			this.button22.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button22.Location = new System.Drawing.Point(768, 184);
			this.button22.Name = "button22";
			this.button22.Size = new System.Drawing.Size(152, 88);
			this.button22.TabIndex = 84;
			this.button22.Text = "Habitacion # 1";
			this.button22.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button22.Click += new System.EventHandler(this.button22_Click);
			this.button22.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button22_MouseDown);
			// 
			// labH69
			// 
			this.labH69.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH69.Location = new System.Drawing.Point(936, 248);
			this.labH69.Name = "labH69";
			this.labH69.Size = new System.Drawing.Size(120, 16);
			this.labH69.TabIndex = 69;
			this.labH69.Text = "Estado";
			this.labH69.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH68
			// 
			this.labH68.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH68.Location = new System.Drawing.Point(1000, 224);
			this.labH68.Name = "labH68";
			this.labH68.Size = new System.Drawing.Size(48, 16);
			this.labH68.TabIndex = 68;
			this.labH68.Text = "Hora 2";
			this.labH68.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH67
			// 
			this.labH67.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH67.Location = new System.Drawing.Point(944, 224);
			this.labH67.Name = "labH67";
			this.labH67.Size = new System.Drawing.Size(48, 16);
			this.labH67.TabIndex = 67;
			this.labH67.Text = "Hora 1";
			this.labH67.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button23
			// 
			this.button23.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button23.Location = new System.Drawing.Point(920, 184);
			this.button23.Name = "button23";
			this.button23.Size = new System.Drawing.Size(152, 88);
			this.button23.TabIndex = 88;
			this.button23.Text = "Habitacion # 1";
			this.button23.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button23.Click += new System.EventHandler(this.button23_Click);
			this.button23.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button23_MouseDown);
			// 
			// labH72
			// 
			this.labH72.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH72.Location = new System.Drawing.Point(1088, 248);
			this.labH72.Name = "labH72";
			this.labH72.Size = new System.Drawing.Size(120, 16);
			this.labH72.TabIndex = 72;
			this.labH72.Text = "Estado";
			this.labH72.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH71
			// 
			this.labH71.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH71.Location = new System.Drawing.Point(1152, 224);
			this.labH71.Name = "labH71";
			this.labH71.Size = new System.Drawing.Size(48, 16);
			this.labH71.TabIndex = 71;
			this.labH71.Text = "Hora 2";
			this.labH71.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH70
			// 
			this.labH70.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH70.Location = new System.Drawing.Point(1096, 224);
			this.labH70.Name = "labH70";
			this.labH70.Size = new System.Drawing.Size(48, 16);
			this.labH70.TabIndex = 70;
			this.labH70.Text = "Hora 1";
			this.labH70.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button24
			// 
			this.button24.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button24.Location = new System.Drawing.Point(1072, 184);
			this.button24.Name = "button24";
			this.button24.Size = new System.Drawing.Size(152, 88);
			this.button24.TabIndex = 92;
			this.button24.Text = "Habitacion # 1";
			this.button24.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button24.Click += new System.EventHandler(this.button24_Click);
			this.button24.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button24_MouseDown);
			// 
			// labH75
			// 
			this.labH75.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH75.Location = new System.Drawing.Point(24, 336);
			this.labH75.Name = "labH75";
			this.labH75.Size = new System.Drawing.Size(120, 16);
			this.labH75.TabIndex = 99;
			this.labH75.Text = "Estado";
			this.labH75.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH74
			// 
			this.labH74.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH74.Location = new System.Drawing.Point(88, 312);
			this.labH74.Name = "labH74";
			this.labH74.Size = new System.Drawing.Size(48, 16);
			this.labH74.TabIndex = 98;
			this.labH74.Text = "Hora 2";
			this.labH74.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH73
			// 
			this.labH73.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH73.Location = new System.Drawing.Point(32, 312);
			this.labH73.Name = "labH73";
			this.labH73.Size = new System.Drawing.Size(48, 16);
			this.labH73.TabIndex = 97;
			this.labH73.Text = "Hora 1";
			this.labH73.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button25
			// 
			this.button25.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button25.Location = new System.Drawing.Point(8, 272);
			this.button25.Name = "button25";
			this.button25.Size = new System.Drawing.Size(152, 88);
			this.button25.TabIndex = 96;
			this.button25.Text = "Habitacion # 1";
			this.button25.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button25.Click += new System.EventHandler(this.button25_Click);
			this.button25.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button25_MouseDown);
			// 
			// labH78
			// 
			this.labH78.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH78.Location = new System.Drawing.Point(176, 336);
			this.labH78.Name = "labH78";
			this.labH78.Size = new System.Drawing.Size(120, 16);
			this.labH78.TabIndex = 103;
			this.labH78.Text = "Estado";
			this.labH78.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH77
			// 
			this.labH77.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH77.Location = new System.Drawing.Point(240, 312);
			this.labH77.Name = "labH77";
			this.labH77.Size = new System.Drawing.Size(48, 16);
			this.labH77.TabIndex = 102;
			this.labH77.Text = "Hora 2";
			this.labH77.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH76
			// 
			this.labH76.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH76.Location = new System.Drawing.Point(184, 312);
			this.labH76.Name = "labH76";
			this.labH76.Size = new System.Drawing.Size(48, 16);
			this.labH76.TabIndex = 101;
			this.labH76.Text = "Hora 1";
			this.labH76.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button26
			// 
			this.button26.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button26.Location = new System.Drawing.Point(160, 272);
			this.button26.Name = "button26";
			this.button26.Size = new System.Drawing.Size(152, 88);
			this.button26.TabIndex = 100;
			this.button26.Text = "Habitacion # 1";
			this.button26.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button26.Click += new System.EventHandler(this.button26_Click);
			this.button26.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button26_MouseDown);
			// 
			// labH81
			// 
			this.labH81.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH81.Location = new System.Drawing.Point(328, 336);
			this.labH81.Name = "labH81";
			this.labH81.Size = new System.Drawing.Size(120, 16);
			this.labH81.TabIndex = 107;
			this.labH81.Text = "Estado";
			this.labH81.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH80
			// 
			this.labH80.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH80.Location = new System.Drawing.Point(392, 312);
			this.labH80.Name = "labH80";
			this.labH80.Size = new System.Drawing.Size(48, 16);
			this.labH80.TabIndex = 106;
			this.labH80.Text = "Hora 2";
			this.labH80.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH79
			// 
			this.labH79.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH79.Location = new System.Drawing.Point(336, 312);
			this.labH79.Name = "labH79";
			this.labH79.Size = new System.Drawing.Size(48, 16);
			this.labH79.TabIndex = 105;
			this.labH79.Text = "Hora 1";
			this.labH79.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button27
			// 
			this.button27.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button27.Location = new System.Drawing.Point(312, 272);
			this.button27.Name = "button27";
			this.button27.Size = new System.Drawing.Size(152, 88);
			this.button27.TabIndex = 104;
			this.button27.Text = "Habitacion # 1";
			this.button27.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button27.Click += new System.EventHandler(this.button27_Click);
			this.button27.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button27_MouseDown);
			// 
			// labH84
			// 
			this.labH84.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH84.Location = new System.Drawing.Point(480, 336);
			this.labH84.Name = "labH84";
			this.labH84.Size = new System.Drawing.Size(120, 16);
			this.labH84.TabIndex = 111;
			this.labH84.Text = "Estado";
			this.labH84.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH83
			// 
			this.labH83.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH83.Location = new System.Drawing.Point(544, 312);
			this.labH83.Name = "labH83";
			this.labH83.Size = new System.Drawing.Size(48, 16);
			this.labH83.TabIndex = 110;
			this.labH83.Text = "Hora 2";
			this.labH83.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH82
			// 
			this.labH82.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH82.Location = new System.Drawing.Point(488, 312);
			this.labH82.Name = "labH82";
			this.labH82.Size = new System.Drawing.Size(48, 16);
			this.labH82.TabIndex = 109;
			this.labH82.Text = "Hora 1";
			this.labH82.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button28
			// 
			this.button28.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button28.Location = new System.Drawing.Point(464, 272);
			this.button28.Name = "button28";
			this.button28.Size = new System.Drawing.Size(152, 88);
			this.button28.TabIndex = 108;
			this.button28.Text = "Habitacion # 1";
			this.button28.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button28.Click += new System.EventHandler(this.button28_Click);
			this.button28.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button28_MouseDown);
			// 
			// labH87
			// 
			this.labH87.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH87.Location = new System.Drawing.Point(632, 336);
			this.labH87.Name = "labH87";
			this.labH87.Size = new System.Drawing.Size(120, 16);
			this.labH87.TabIndex = 115;
			this.labH87.Text = "Estado";
			this.labH87.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH86
			// 
			this.labH86.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH86.Location = new System.Drawing.Point(696, 312);
			this.labH86.Name = "labH86";
			this.labH86.Size = new System.Drawing.Size(48, 16);
			this.labH86.TabIndex = 114;
			this.labH86.Text = "Hora 2";
			this.labH86.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH85
			// 
			this.labH85.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH85.Location = new System.Drawing.Point(640, 312);
			this.labH85.Name = "labH85";
			this.labH85.Size = new System.Drawing.Size(48, 16);
			this.labH85.TabIndex = 113;
			this.labH85.Text = "Hora 1";
			this.labH85.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button29
			// 
			this.button29.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button29.Location = new System.Drawing.Point(616, 272);
			this.button29.Name = "button29";
			this.button29.Size = new System.Drawing.Size(152, 88);
			this.button29.TabIndex = 112;
			this.button29.Text = "Habitacion # 1";
			this.button29.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button29.Click += new System.EventHandler(this.button29_Click);
			this.button29.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button29_MouseDown);
			// 
			// labH90
			// 
			this.labH90.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH90.Location = new System.Drawing.Point(784, 336);
			this.labH90.Name = "labH90";
			this.labH90.Size = new System.Drawing.Size(120, 16);
			this.labH90.TabIndex = 119;
			this.labH90.Text = "Estado";
			this.labH90.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH89
			// 
			this.labH89.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH89.Location = new System.Drawing.Point(848, 312);
			this.labH89.Name = "labH89";
			this.labH89.Size = new System.Drawing.Size(48, 16);
			this.labH89.TabIndex = 118;
			this.labH89.Text = "Hora 2";
			this.labH89.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH88
			// 
			this.labH88.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH88.Location = new System.Drawing.Point(792, 312);
			this.labH88.Name = "labH88";
			this.labH88.Size = new System.Drawing.Size(48, 16);
			this.labH88.TabIndex = 117;
			this.labH88.Text = "Hora 1";
			this.labH88.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button30
			// 
			this.button30.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button30.Location = new System.Drawing.Point(768, 272);
			this.button30.Name = "button30";
			this.button30.Size = new System.Drawing.Size(152, 88);
			this.button30.TabIndex = 116;
			this.button30.Text = "Habitacion # 1";
			this.button30.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button30.Click += new System.EventHandler(this.button30_Click);
			this.button30.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button30_MouseDown);
			// 
			// labH93
			// 
			this.labH93.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH93.Location = new System.Drawing.Point(936, 336);
			this.labH93.Name = "labH93";
			this.labH93.Size = new System.Drawing.Size(120, 16);
			this.labH93.TabIndex = 123;
			this.labH93.Text = "Estado";
			this.labH93.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH92
			// 
			this.labH92.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH92.Location = new System.Drawing.Point(1000, 312);
			this.labH92.Name = "labH92";
			this.labH92.Size = new System.Drawing.Size(48, 16);
			this.labH92.TabIndex = 122;
			this.labH92.Text = "Hora 2";
			this.labH92.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH91
			// 
			this.labH91.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH91.Location = new System.Drawing.Point(944, 312);
			this.labH91.Name = "labH91";
			this.labH91.Size = new System.Drawing.Size(48, 16);
			this.labH91.TabIndex = 121;
			this.labH91.Text = "Hora 1";
			this.labH91.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button31
			// 
			this.button31.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button31.Location = new System.Drawing.Point(920, 272);
			this.button31.Name = "button31";
			this.button31.Size = new System.Drawing.Size(152, 88);
			this.button31.TabIndex = 120;
			this.button31.Text = "Habitacion # 1";
			this.button31.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button31.Click += new System.EventHandler(this.button31_Click);
			this.button31.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button31_MouseDown);
			// 
			// labH96
			// 
			this.labH96.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH96.Location = new System.Drawing.Point(1088, 336);
			this.labH96.Name = "labH96";
			this.labH96.Size = new System.Drawing.Size(120, 16);
			this.labH96.TabIndex = 127;
			this.labH96.Text = "Estado";
			this.labH96.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH95
			// 
			this.labH95.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH95.Location = new System.Drawing.Point(1152, 312);
			this.labH95.Name = "labH95";
			this.labH95.Size = new System.Drawing.Size(48, 16);
			this.labH95.TabIndex = 126;
			this.labH95.Text = "Hora 2";
			this.labH95.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH94
			// 
			this.labH94.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH94.Location = new System.Drawing.Point(1096, 312);
			this.labH94.Name = "labH94";
			this.labH94.Size = new System.Drawing.Size(48, 16);
			this.labH94.TabIndex = 125;
			this.labH94.Text = "Hora 1";
			this.labH94.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button32
			// 
			this.button32.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button32.Location = new System.Drawing.Point(1072, 272);
			this.button32.Name = "button32";
			this.button32.Size = new System.Drawing.Size(152, 88);
			this.button32.TabIndex = 124;
			this.button32.Text = "Habitacion # 1";
			this.button32.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button32.Click += new System.EventHandler(this.button32_Click);
			this.button32.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button32_MouseDown);
			// 
			// labH99
			// 
			this.labH99.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH99.Location = new System.Drawing.Point(24, 424);
			this.labH99.Name = "labH99";
			this.labH99.Size = new System.Drawing.Size(120, 16);
			this.labH99.TabIndex = 75;
			this.labH99.Text = "Estado";
			this.labH99.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH98
			// 
			this.labH98.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH98.Location = new System.Drawing.Point(88, 400);
			this.labH98.Name = "labH98";
			this.labH98.Size = new System.Drawing.Size(48, 16);
			this.labH98.TabIndex = 74;
			this.labH98.Text = "Hora 2";
			this.labH98.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH97
			// 
			this.labH97.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH97.Location = new System.Drawing.Point(32, 400);
			this.labH97.Name = "labH97";
			this.labH97.Size = new System.Drawing.Size(48, 16);
			this.labH97.TabIndex = 73;
			this.labH97.Text = "Hora 1";
			this.labH97.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button33
			// 
			this.button33.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button33.Location = new System.Drawing.Point(8, 360);
			this.button33.Name = "button33";
			this.button33.Size = new System.Drawing.Size(152, 88);
			this.button33.TabIndex = 128;
			this.button33.Text = "Habitacion # 1";
			this.button33.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button33.Click += new System.EventHandler(this.button33_Click);
			this.button33.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button33_MouseDown);
			// 
			// labH102
			// 
			this.labH102.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH102.Location = new System.Drawing.Point(176, 424);
			this.labH102.Name = "labH102";
			this.labH102.Size = new System.Drawing.Size(120, 16);
			this.labH102.TabIndex = 78;
			this.labH102.Text = "Estado";
			this.labH102.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH101
			// 
			this.labH101.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH101.Location = new System.Drawing.Point(240, 400);
			this.labH101.Name = "labH101";
			this.labH101.Size = new System.Drawing.Size(48, 16);
			this.labH101.TabIndex = 77;
			this.labH101.Text = "Hora 2";
			this.labH101.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH100
			// 
			this.labH100.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH100.Location = new System.Drawing.Point(184, 400);
			this.labH100.Name = "labH100";
			this.labH100.Size = new System.Drawing.Size(48, 16);
			this.labH100.TabIndex = 76;
			this.labH100.Text = "Hora 1";
			this.labH100.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button34
			// 
			this.button34.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button34.Location = new System.Drawing.Point(160, 360);
			this.button34.Name = "button34";
			this.button34.Size = new System.Drawing.Size(152, 88);
			this.button34.TabIndex = 132;
			this.button34.Text = "Habitacion # 1";
			this.button34.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button34.Click += new System.EventHandler(this.button34_Click);
			this.button34.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button34_MouseDown);
			// 
			// labH105
			// 
			this.labH105.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH105.Location = new System.Drawing.Point(328, 424);
			this.labH105.Name = "labH105";
			this.labH105.Size = new System.Drawing.Size(120, 16);
			this.labH105.TabIndex = 81;
			this.labH105.Text = "Estado";
			this.labH105.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH104
			// 
			this.labH104.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH104.Location = new System.Drawing.Point(392, 400);
			this.labH104.Name = "labH104";
			this.labH104.Size = new System.Drawing.Size(48, 16);
			this.labH104.TabIndex = 80;
			this.labH104.Text = "Hora 2";
			this.labH104.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH103
			// 
			this.labH103.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH103.Location = new System.Drawing.Point(336, 400);
			this.labH103.Name = "labH103";
			this.labH103.Size = new System.Drawing.Size(48, 16);
			this.labH103.TabIndex = 79;
			this.labH103.Text = "Hora 1";
			this.labH103.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button35
			// 
			this.button35.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button35.Location = new System.Drawing.Point(312, 360);
			this.button35.Name = "button35";
			this.button35.Size = new System.Drawing.Size(152, 88);
			this.button35.TabIndex = 136;
			this.button35.Text = "Habitacion # 1";
			this.button35.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button35.Click += new System.EventHandler(this.button35_Click);
			this.button35.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button35_MouseDown);
			// 
			// labH108
			// 
			this.labH108.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH108.Location = new System.Drawing.Point(480, 424);
			this.labH108.Name = "labH108";
			this.labH108.Size = new System.Drawing.Size(120, 16);
			this.labH108.TabIndex = 84;
			this.labH108.Text = "Estado";
			this.labH108.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH107
			// 
			this.labH107.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH107.Location = new System.Drawing.Point(544, 400);
			this.labH107.Name = "labH107";
			this.labH107.Size = new System.Drawing.Size(48, 16);
			this.labH107.TabIndex = 83;
			this.labH107.Text = "Hora 2";
			this.labH107.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH106
			// 
			this.labH106.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH106.Location = new System.Drawing.Point(488, 400);
			this.labH106.Name = "labH106";
			this.labH106.Size = new System.Drawing.Size(48, 16);
			this.labH106.TabIndex = 82;
			this.labH106.Text = "Hora 1";
			this.labH106.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button36
			// 
			this.button36.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button36.Location = new System.Drawing.Point(464, 360);
			this.button36.Name = "button36";
			this.button36.Size = new System.Drawing.Size(152, 88);
			this.button36.TabIndex = 140;
			this.button36.Text = "Habitacion # 1";
			this.button36.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button36.Click += new System.EventHandler(this.button36_Click);
			this.button36.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button36_MouseDown);
			// 
			// labH111
			// 
			this.labH111.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH111.Location = new System.Drawing.Point(632, 424);
			this.labH111.Name = "labH111";
			this.labH111.Size = new System.Drawing.Size(120, 16);
			this.labH111.TabIndex = 87;
			this.labH111.Text = "Estado";
			this.labH111.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH110
			// 
			this.labH110.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH110.Location = new System.Drawing.Point(696, 400);
			this.labH110.Name = "labH110";
			this.labH110.Size = new System.Drawing.Size(48, 16);
			this.labH110.TabIndex = 86;
			this.labH110.Text = "Hora 2";
			this.labH110.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH109
			// 
			this.labH109.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH109.Location = new System.Drawing.Point(640, 400);
			this.labH109.Name = "labH109";
			this.labH109.Size = new System.Drawing.Size(48, 16);
			this.labH109.TabIndex = 85;
			this.labH109.Text = "Hora 1";
			this.labH109.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button37
			// 
			this.button37.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button37.Location = new System.Drawing.Point(616, 360);
			this.button37.Name = "button37";
			this.button37.Size = new System.Drawing.Size(152, 88);
			this.button37.TabIndex = 144;
			this.button37.Text = "Habitacion # 1";
			this.button37.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button37.Click += new System.EventHandler(this.button37_Click);
			this.button37.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button37_MouseDown);
			// 
			// labH114
			// 
			this.labH114.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH114.Location = new System.Drawing.Point(784, 424);
			this.labH114.Name = "labH114";
			this.labH114.Size = new System.Drawing.Size(120, 16);
			this.labH114.TabIndex = 90;
			this.labH114.Text = "Estado";
			this.labH114.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH113
			// 
			this.labH113.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH113.Location = new System.Drawing.Point(848, 400);
			this.labH113.Name = "labH113";
			this.labH113.Size = new System.Drawing.Size(48, 16);
			this.labH113.TabIndex = 89;
			this.labH113.Text = "Hora 2";
			this.labH113.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH112
			// 
			this.labH112.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH112.Location = new System.Drawing.Point(792, 400);
			this.labH112.Name = "labH112";
			this.labH112.Size = new System.Drawing.Size(48, 16);
			this.labH112.TabIndex = 88;
			this.labH112.Text = "Hora 1";
			this.labH112.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button38
			// 
			this.button38.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button38.Location = new System.Drawing.Point(768, 360);
			this.button38.Name = "button38";
			this.button38.Size = new System.Drawing.Size(152, 88);
			this.button38.TabIndex = 148;
			this.button38.Text = "Habitacion # 1";
			this.button38.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button38.Click += new System.EventHandler(this.button38_Click);
			this.button38.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button38_MouseDown);
			// 
			// labH117
			// 
			this.labH117.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH117.Location = new System.Drawing.Point(936, 424);
			this.labH117.Name = "labH117";
			this.labH117.Size = new System.Drawing.Size(120, 16);
			this.labH117.TabIndex = 93;
			this.labH117.Text = "Estado";
			this.labH117.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH116
			// 
			this.labH116.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH116.Location = new System.Drawing.Point(1000, 400);
			this.labH116.Name = "labH116";
			this.labH116.Size = new System.Drawing.Size(48, 16);
			this.labH116.TabIndex = 92;
			this.labH116.Text = "Hora 2";
			this.labH116.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH115
			// 
			this.labH115.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH115.Location = new System.Drawing.Point(944, 400);
			this.labH115.Name = "labH115";
			this.labH115.Size = new System.Drawing.Size(48, 16);
			this.labH115.TabIndex = 91;
			this.labH115.Text = "Hora 1";
			this.labH115.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button39
			// 
			this.button39.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button39.Location = new System.Drawing.Point(920, 360);
			this.button39.Name = "button39";
			this.button39.Size = new System.Drawing.Size(152, 88);
			this.button39.TabIndex = 152;
			this.button39.Text = "Habitacion # 1";
			this.button39.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button39.Click += new System.EventHandler(this.button39_Click);
			this.button39.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button39_MouseDown);
			// 
			// labH120
			// 
			this.labH120.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH120.Location = new System.Drawing.Point(1088, 424);
			this.labH120.Name = "labH120";
			this.labH120.Size = new System.Drawing.Size(120, 16);
			this.labH120.TabIndex = 96;
			this.labH120.Text = "Estado";
			this.labH120.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH119
			// 
			this.labH119.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH119.Location = new System.Drawing.Point(1152, 400);
			this.labH119.Name = "labH119";
			this.labH119.Size = new System.Drawing.Size(48, 16);
			this.labH119.TabIndex = 95;
			this.labH119.Text = "Hora 2";
			this.labH119.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH118
			// 
			this.labH118.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH118.Location = new System.Drawing.Point(1096, 400);
			this.labH118.Name = "labH118";
			this.labH118.Size = new System.Drawing.Size(48, 16);
			this.labH118.TabIndex = 94;
			this.labH118.Text = "Hora 1";
			this.labH118.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button40
			// 
			this.button40.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button40.Location = new System.Drawing.Point(1072, 360);
			this.button40.Name = "button40";
			this.button40.Size = new System.Drawing.Size(152, 88);
			this.button40.TabIndex = 156;
			this.button40.Text = "Habitacion # 1";
			this.button40.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button40.Click += new System.EventHandler(this.button40_Click);
			this.button40.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button40_MouseDown);
			// 
			// labH123
			// 
			this.labH123.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH123.Location = new System.Drawing.Point(24, 512);
			this.labH123.Name = "labH123";
			this.labH123.Size = new System.Drawing.Size(120, 16);
			this.labH123.TabIndex = 99;
			this.labH123.Text = "Estado";
			this.labH123.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH122
			// 
			this.labH122.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH122.Location = new System.Drawing.Point(88, 488);
			this.labH122.Name = "labH122";
			this.labH122.Size = new System.Drawing.Size(48, 16);
			this.labH122.TabIndex = 98;
			this.labH122.Text = "Hora 2";
			this.labH122.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH121
			// 
			this.labH121.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH121.Location = new System.Drawing.Point(32, 488);
			this.labH121.Name = "labH121";
			this.labH121.Size = new System.Drawing.Size(48, 16);
			this.labH121.TabIndex = 97;
			this.labH121.Text = "Hora 1";
			this.labH121.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button41
			// 
			this.button41.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button41.Location = new System.Drawing.Point(8, 448);
			this.button41.Name = "button41";
			this.button41.Size = new System.Drawing.Size(152, 88);
			this.button41.TabIndex = 160;
			this.button41.Text = "Habitacion # 1";
			this.button41.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button41.Click += new System.EventHandler(this.button41_Click);
			this.button41.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button41_MouseDown);
			// 
			// labH126
			// 
			this.labH126.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH126.Location = new System.Drawing.Point(176, 512);
			this.labH126.Name = "labH126";
			this.labH126.Size = new System.Drawing.Size(120, 16);
			this.labH126.TabIndex = 102;
			this.labH126.Text = "Estado";
			this.labH126.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH125
			// 
			this.labH125.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH125.Location = new System.Drawing.Point(240, 488);
			this.labH125.Name = "labH125";
			this.labH125.Size = new System.Drawing.Size(48, 16);
			this.labH125.TabIndex = 101;
			this.labH125.Text = "Hora 2";
			this.labH125.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH124
			// 
			this.labH124.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH124.Location = new System.Drawing.Point(184, 488);
			this.labH124.Name = "labH124";
			this.labH124.Size = new System.Drawing.Size(48, 16);
			this.labH124.TabIndex = 100;
			this.labH124.Text = "Hora 1";
			this.labH124.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button42
			// 
			this.button42.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button42.Location = new System.Drawing.Point(160, 448);
			this.button42.Name = "button42";
			this.button42.Size = new System.Drawing.Size(152, 88);
			this.button42.TabIndex = 164;
			this.button42.Text = "Habitacion # 1";
			this.button42.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button42.Click += new System.EventHandler(this.button42_Click);
			this.button42.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button42_MouseDown);
			// 
			// labH129
			// 
			this.labH129.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH129.Location = new System.Drawing.Point(328, 512);
			this.labH129.Name = "labH129";
			this.labH129.Size = new System.Drawing.Size(120, 16);
			this.labH129.TabIndex = 105;
			this.labH129.Text = "Estado";
			this.labH129.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH128
			// 
			this.labH128.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH128.Location = new System.Drawing.Point(392, 488);
			this.labH128.Name = "labH128";
			this.labH128.Size = new System.Drawing.Size(48, 16);
			this.labH128.TabIndex = 104;
			this.labH128.Text = "Hora 2";
			this.labH128.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH127
			// 
			this.labH127.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH127.Location = new System.Drawing.Point(336, 488);
			this.labH127.Name = "labH127";
			this.labH127.Size = new System.Drawing.Size(48, 16);
			this.labH127.TabIndex = 103;
			this.labH127.Text = "Hora 1";
			this.labH127.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button43
			// 
			this.button43.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button43.Location = new System.Drawing.Point(312, 448);
			this.button43.Name = "button43";
			this.button43.Size = new System.Drawing.Size(152, 88);
			this.button43.TabIndex = 168;
			this.button43.Text = "Habitacion # 1";
			this.button43.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button43.Click += new System.EventHandler(this.button43_Click);
			this.button43.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button43_MouseDown);
			// 
			// labH132
			// 
			this.labH132.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH132.Location = new System.Drawing.Point(480, 512);
			this.labH132.Name = "labH132";
			this.labH132.Size = new System.Drawing.Size(120, 16);
			this.labH132.TabIndex = 108;
			this.labH132.Text = "Estado";
			this.labH132.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH131
			// 
			this.labH131.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH131.Location = new System.Drawing.Point(544, 488);
			this.labH131.Name = "labH131";
			this.labH131.Size = new System.Drawing.Size(48, 16);
			this.labH131.TabIndex = 107;
			this.labH131.Text = "Hora 2";
			this.labH131.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH130
			// 
			this.labH130.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH130.Location = new System.Drawing.Point(488, 488);
			this.labH130.Name = "labH130";
			this.labH130.Size = new System.Drawing.Size(48, 16);
			this.labH130.TabIndex = 106;
			this.labH130.Text = "Hora 1";
			this.labH130.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button44
			// 
			this.button44.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button44.Location = new System.Drawing.Point(464, 448);
			this.button44.Name = "button44";
			this.button44.Size = new System.Drawing.Size(152, 88);
			this.button44.TabIndex = 172;
			this.button44.Text = "Habitacion # 1";
			this.button44.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button44.Click += new System.EventHandler(this.button44_Click);
			this.button44.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button44_MouseDown);
			// 
			// labH135
			// 
			this.labH135.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH135.Location = new System.Drawing.Point(632, 512);
			this.labH135.Name = "labH135";
			this.labH135.Size = new System.Drawing.Size(120, 16);
			this.labH135.TabIndex = 111;
			this.labH135.Text = "Estado";
			this.labH135.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH134
			// 
			this.labH134.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH134.Location = new System.Drawing.Point(696, 488);
			this.labH134.Name = "labH134";
			this.labH134.Size = new System.Drawing.Size(48, 16);
			this.labH134.TabIndex = 110;
			this.labH134.Text = "Hora 2";
			this.labH134.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH133
			// 
			this.labH133.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH133.Location = new System.Drawing.Point(640, 488);
			this.labH133.Name = "labH133";
			this.labH133.Size = new System.Drawing.Size(48, 16);
			this.labH133.TabIndex = 108;
			this.labH133.Text = "Hora 1";
			this.labH133.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button45
			// 
			this.button45.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button45.Location = new System.Drawing.Point(616, 448);
			this.button45.Name = "button45";
			this.button45.Size = new System.Drawing.Size(152, 88);
			this.button45.TabIndex = 176;
			this.button45.Text = "Habitacion # 1";
			this.button45.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button45.Click += new System.EventHandler(this.button45_Click);
			this.button45.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button45_MouseDown);
			// 
			// labH138
			// 
			this.labH138.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH138.Location = new System.Drawing.Point(784, 512);
			this.labH138.Name = "labH138";
			this.labH138.Size = new System.Drawing.Size(120, 16);
			this.labH138.TabIndex = 114;
			this.labH138.Text = "Estado";
			this.labH138.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH137
			// 
			this.labH137.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH137.Location = new System.Drawing.Point(848, 488);
			this.labH137.Name = "labH137";
			this.labH137.Size = new System.Drawing.Size(48, 16);
			this.labH137.TabIndex = 113;
			this.labH137.Text = "Hora 2";
			this.labH137.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH136
			// 
			this.labH136.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH136.Location = new System.Drawing.Point(792, 488);
			this.labH136.Name = "labH136";
			this.labH136.Size = new System.Drawing.Size(48, 16);
			this.labH136.TabIndex = 112;
			this.labH136.Text = "Hora 1";
			this.labH136.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button46
			// 
			this.button46.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button46.Location = new System.Drawing.Point(768, 448);
			this.button46.Name = "button46";
			this.button46.Size = new System.Drawing.Size(152, 88);
			this.button46.TabIndex = 180;
			this.button46.Text = "Habitacion # 1";
			this.button46.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button46.Click += new System.EventHandler(this.button46_Click);
			this.button46.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button46_MouseDown);
			// 
			// labH141
			// 
			this.labH141.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH141.Location = new System.Drawing.Point(936, 512);
			this.labH141.Name = "labH141";
			this.labH141.Size = new System.Drawing.Size(120, 16);
			this.labH141.TabIndex = 117;
			this.labH141.Text = "Estado";
			this.labH141.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH140
			// 
			this.labH140.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH140.Location = new System.Drawing.Point(1000, 488);
			this.labH140.Name = "labH140";
			this.labH140.Size = new System.Drawing.Size(48, 16);
			this.labH140.TabIndex = 116;
			this.labH140.Text = "Hora 2";
			this.labH140.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH139
			// 
			this.labH139.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH139.Location = new System.Drawing.Point(944, 488);
			this.labH139.Name = "labH139";
			this.labH139.Size = new System.Drawing.Size(48, 16);
			this.labH139.TabIndex = 115;
			this.labH139.Text = "Hora 1";
			this.labH139.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button47
			// 
			this.button47.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button47.Location = new System.Drawing.Point(920, 448);
			this.button47.Name = "button47";
			this.button47.Size = new System.Drawing.Size(152, 88);
			this.button47.TabIndex = 184;
			this.button47.Text = "Habitacion # 1";
			this.button47.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button47.Click += new System.EventHandler(this.button47_Click);
			this.button47.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button47_MouseDown);
			// 
			// labH144
			// 
			this.labH144.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH144.Location = new System.Drawing.Point(1088, 512);
			this.labH144.Name = "labH144";
			this.labH144.Size = new System.Drawing.Size(120, 16);
			this.labH144.TabIndex = 120;
			this.labH144.Text = "Estado";
			this.labH144.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH143
			// 
			this.labH143.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH143.Location = new System.Drawing.Point(1152, 488);
			this.labH143.Name = "labH143";
			this.labH143.Size = new System.Drawing.Size(48, 16);
			this.labH143.TabIndex = 119;
			this.labH143.Text = "Hora 2";
			this.labH143.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// labH142
			// 
			this.labH142.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labH142.Location = new System.Drawing.Point(1096, 488);
			this.labH142.Name = "labH142";
			this.labH142.Size = new System.Drawing.Size(48, 16);
			this.labH142.TabIndex = 118;
			this.labH142.Text = "Hora 1";
			this.labH142.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// button48
			// 
			this.button48.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button48.Location = new System.Drawing.Point(1072, 448);
			this.button48.Name = "button48";
			this.button48.Size = new System.Drawing.Size(152, 88);
			this.button48.TabIndex = 188;
			this.button48.Text = "Habitacion # 1";
			this.button48.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.button48.Click += new System.EventHandler(this.button48_Click);
			this.button48.MouseDown += new System.Windows.Forms.MouseEventHandler(this.button48_MouseDown);
			// 
			// timer1
			// 
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// MnuEstado
			// 
			this.MnuEstado.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							this.mnuLimpieza,
																																							this.mnuLibre,
																																							this.menuIngresoCLI,
																																							this.menuEnvioCB});
			// 
			// mnuLimpieza
			// 
			this.mnuLimpieza.Index = 0;
			this.mnuLimpieza.Text = "Limpieza";
			this.mnuLimpieza.Click += new System.EventHandler(this.mnuLimpieza_Click);
			// 
			// mnuLibre
			// 
			this.mnuLibre.Index = 1;
			this.mnuLibre.Text = "Libre";
			this.mnuLibre.Click += new System.EventHandler(this.mnuLibre_Click);
			// 
			// menuIngresoCLI
			// 
			this.menuIngresoCLI.Index = 2;
			this.menuIngresoCLI.Text = "Ingreso Cliente";
			this.menuIngresoCLI.Click += new System.EventHandler(this.menuIngresoCLI_Click);
			// 
			// menuEnvioCB
			// 
			this.menuEnvioCB.Index = 3;
			this.menuEnvioCB.Text = "Envio Cobro";
			this.menuEnvioCB.Click += new System.EventHandler(this.menuEnvioCB_Click);
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
			// dtFecha
			// 
			this.dtFecha.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.FontData.BoldAsString = "True";
			appearance1.FontData.Name = "Tahoma";
			appearance1.FontData.SizeInPoints = 10F;
			appearance1.ForeColor = System.Drawing.Color.Firebrick;
			appearance1.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy HH:mm";
			this.dtFecha.Location = new System.Drawing.Point(8, 512);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(248, 22);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 701;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.Visible = false;
			// 
			// timer2
			// 
			this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
			// 
			// labelEn1
			// 
			this.labelEn1.BackColor = System.Drawing.Color.Green;
			this.labelEn1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn1.ForeColor = System.Drawing.Color.White;
			this.labelEn1.Location = new System.Drawing.Point(24, 32);
			this.labelEn1.Name = "labelEn1";
			this.labelEn1.Size = new System.Drawing.Size(120, 16);
			this.labelEn1.TabIndex = 702;
			this.labelEn1.Text = "Estado";
			this.labelEn1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn1.Visible = false;
			// 
			// labelEn2
			// 
			this.labelEn2.BackColor = System.Drawing.Color.Green;
			this.labelEn2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn2.ForeColor = System.Drawing.Color.White;
			this.labelEn2.Location = new System.Drawing.Point(176, 32);
			this.labelEn2.Name = "labelEn2";
			this.labelEn2.Size = new System.Drawing.Size(120, 16);
			this.labelEn2.TabIndex = 703;
			this.labelEn2.Text = "Estado";
			this.labelEn2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn2.Visible = false;
			// 
			// labelEn3
			// 
			this.labelEn3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn3.Location = new System.Drawing.Point(328, 32);
			this.labelEn3.Name = "labelEn3";
			this.labelEn3.Size = new System.Drawing.Size(120, 16);
			this.labelEn3.TabIndex = 704;
			this.labelEn3.Text = "Estado";
			this.labelEn3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn3.Visible = false;
			// 
			// labelEn4
			// 
			this.labelEn4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn4.Location = new System.Drawing.Point(480, 32);
			this.labelEn4.Name = "labelEn4";
			this.labelEn4.Size = new System.Drawing.Size(120, 16);
			this.labelEn4.TabIndex = 705;
			this.labelEn4.Text = "Estado";
			this.labelEn4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn4.Visible = false;
			// 
			// labelEn5
			// 
			this.labelEn5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn5.Location = new System.Drawing.Point(632, 32);
			this.labelEn5.Name = "labelEn5";
			this.labelEn5.Size = new System.Drawing.Size(120, 16);
			this.labelEn5.TabIndex = 706;
			this.labelEn5.Text = "Estado";
			this.labelEn5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn5.Visible = false;
			// 
			// labelEn6
			// 
			this.labelEn6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn6.Location = new System.Drawing.Point(784, 32);
			this.labelEn6.Name = "labelEn6";
			this.labelEn6.Size = new System.Drawing.Size(120, 16);
			this.labelEn6.TabIndex = 707;
			this.labelEn6.Text = "Estado";
			this.labelEn6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn6.Visible = false;
			// 
			// labelEn29
			// 
			this.labelEn29.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn29.Location = new System.Drawing.Point(632, 296);
			this.labelEn29.Name = "labelEn29";
			this.labelEn29.Size = new System.Drawing.Size(120, 16);
			this.labelEn29.TabIndex = 708;
			this.labelEn29.Text = "Estado";
			this.labelEn29.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn29.Visible = false;
			this.labelEn29.Click += new System.EventHandler(this.label7_Click);
			// 
			// labelEn28
			// 
			this.labelEn28.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn28.Location = new System.Drawing.Point(480, 296);
			this.labelEn28.Name = "labelEn28";
			this.labelEn28.Size = new System.Drawing.Size(120, 16);
			this.labelEn28.TabIndex = 709;
			this.labelEn28.Text = "Estado";
			this.labelEn28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn28.Visible = false;
			// 
			// labelEn27
			// 
			this.labelEn27.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn27.Location = new System.Drawing.Point(328, 296);
			this.labelEn27.Name = "labelEn27";
			this.labelEn27.Size = new System.Drawing.Size(120, 16);
			this.labelEn27.TabIndex = 710;
			this.labelEn27.Text = "Estado";
			this.labelEn27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn27.Visible = false;
			// 
			// labelEn26
			// 
			this.labelEn26.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn26.Location = new System.Drawing.Point(176, 296);
			this.labelEn26.Name = "labelEn26";
			this.labelEn26.Size = new System.Drawing.Size(120, 16);
			this.labelEn26.TabIndex = 711;
			this.labelEn26.Text = "Estado";
			this.labelEn26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn26.Visible = false;
			// 
			// labelEn25
			// 
			this.labelEn25.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn25.Location = new System.Drawing.Point(24, 296);
			this.labelEn25.Name = "labelEn25";
			this.labelEn25.Size = new System.Drawing.Size(120, 16);
			this.labelEn25.TabIndex = 712;
			this.labelEn25.Text = "Estado";
			this.labelEn25.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn25.Visible = false;
			// 
			// labelEn24
			// 
			this.labelEn24.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn24.Location = new System.Drawing.Point(1088, 208);
			this.labelEn24.Name = "labelEn24";
			this.labelEn24.Size = new System.Drawing.Size(120, 16);
			this.labelEn24.TabIndex = 713;
			this.labelEn24.Text = "Estado";
			this.labelEn24.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn24.Visible = false;
			// 
			// labelEn23
			// 
			this.labelEn23.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn23.Location = new System.Drawing.Point(936, 208);
			this.labelEn23.Name = "labelEn23";
			this.labelEn23.Size = new System.Drawing.Size(120, 16);
			this.labelEn23.TabIndex = 714;
			this.labelEn23.Text = "Estado";
			this.labelEn23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn23.Visible = false;
			// 
			// labelEn22
			// 
			this.labelEn22.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn22.Location = new System.Drawing.Point(784, 208);
			this.labelEn22.Name = "labelEn22";
			this.labelEn22.Size = new System.Drawing.Size(120, 16);
			this.labelEn22.TabIndex = 715;
			this.labelEn22.Text = "Estado";
			this.labelEn22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn22.Visible = false;
			// 
			// labelEn21
			// 
			this.labelEn21.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn21.Location = new System.Drawing.Point(632, 208);
			this.labelEn21.Name = "labelEn21";
			this.labelEn21.Size = new System.Drawing.Size(120, 16);
			this.labelEn21.TabIndex = 716;
			this.labelEn21.Text = "Estado";
			this.labelEn21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn21.Visible = false;
			// 
			// labelEn20
			// 
			this.labelEn20.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn20.Location = new System.Drawing.Point(480, 208);
			this.labelEn20.Name = "labelEn20";
			this.labelEn20.Size = new System.Drawing.Size(120, 16);
			this.labelEn20.TabIndex = 717;
			this.labelEn20.Text = "Estado";
			this.labelEn20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn20.Visible = false;
			// 
			// labelEn19
			// 
			this.labelEn19.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn19.Location = new System.Drawing.Point(328, 208);
			this.labelEn19.Name = "labelEn19";
			this.labelEn19.Size = new System.Drawing.Size(120, 16);
			this.labelEn19.TabIndex = 718;
			this.labelEn19.Text = "Estado";
			this.labelEn19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn19.Visible = false;
			// 
			// labelEn18
			// 
			this.labelEn18.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn18.Location = new System.Drawing.Point(176, 208);
			this.labelEn18.Name = "labelEn18";
			this.labelEn18.Size = new System.Drawing.Size(120, 16);
			this.labelEn18.TabIndex = 719;
			this.labelEn18.Text = "Estado";
			this.labelEn18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn18.Visible = false;
			// 
			// labelEn17
			// 
			this.labelEn17.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn17.Location = new System.Drawing.Point(24, 208);
			this.labelEn17.Name = "labelEn17";
			this.labelEn17.Size = new System.Drawing.Size(120, 16);
			this.labelEn17.TabIndex = 720;
			this.labelEn17.Text = "Estado";
			this.labelEn17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn17.Visible = false;
			// 
			// labelEn16
			// 
			this.labelEn16.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn16.Location = new System.Drawing.Point(1096, 120);
			this.labelEn16.Name = "labelEn16";
			this.labelEn16.Size = new System.Drawing.Size(120, 16);
			this.labelEn16.TabIndex = 721;
			this.labelEn16.Text = "Estado";
			this.labelEn16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn16.Visible = false;
			// 
			// labelEn15
			// 
			this.labelEn15.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn15.Location = new System.Drawing.Point(936, 120);
			this.labelEn15.Name = "labelEn15";
			this.labelEn15.Size = new System.Drawing.Size(120, 16);
			this.labelEn15.TabIndex = 722;
			this.labelEn15.Text = "Estado";
			this.labelEn15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn15.Visible = false;
			// 
			// labelEn14
			// 
			this.labelEn14.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn14.Location = new System.Drawing.Point(784, 120);
			this.labelEn14.Name = "labelEn14";
			this.labelEn14.Size = new System.Drawing.Size(120, 16);
			this.labelEn14.TabIndex = 723;
			this.labelEn14.Text = "Estado";
			this.labelEn14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn14.Visible = false;
			// 
			// labelEn13
			// 
			this.labelEn13.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn13.Location = new System.Drawing.Point(632, 120);
			this.labelEn13.Name = "labelEn13";
			this.labelEn13.Size = new System.Drawing.Size(120, 16);
			this.labelEn13.TabIndex = 724;
			this.labelEn13.Text = "Estado";
			this.labelEn13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn13.Visible = false;
			// 
			// labelEn12
			// 
			this.labelEn12.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn12.Location = new System.Drawing.Point(480, 120);
			this.labelEn12.Name = "labelEn12";
			this.labelEn12.Size = new System.Drawing.Size(120, 16);
			this.labelEn12.TabIndex = 725;
			this.labelEn12.Text = "Estado";
			this.labelEn12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn12.Visible = false;
			// 
			// labelEn11
			// 
			this.labelEn11.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn11.Location = new System.Drawing.Point(328, 120);
			this.labelEn11.Name = "labelEn11";
			this.labelEn11.Size = new System.Drawing.Size(120, 16);
			this.labelEn11.TabIndex = 726;
			this.labelEn11.Text = "Estado";
			this.labelEn11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn11.Visible = false;
			// 
			// labelEn10
			// 
			this.labelEn10.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn10.Location = new System.Drawing.Point(176, 120);
			this.labelEn10.Name = "labelEn10";
			this.labelEn10.Size = new System.Drawing.Size(120, 16);
			this.labelEn10.TabIndex = 727;
			this.labelEn10.Text = "Estado";
			this.labelEn10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn10.Visible = false;
			// 
			// labelEn9
			// 
			this.labelEn9.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn9.Location = new System.Drawing.Point(24, 120);
			this.labelEn9.Name = "labelEn9";
			this.labelEn9.Size = new System.Drawing.Size(120, 16);
			this.labelEn9.TabIndex = 728;
			this.labelEn9.Text = "Estado";
			this.labelEn9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn9.Visible = false;
			// 
			// labelEn8
			// 
			this.labelEn8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn8.Location = new System.Drawing.Point(1088, 32);
			this.labelEn8.Name = "labelEn8";
			this.labelEn8.Size = new System.Drawing.Size(120, 16);
			this.labelEn8.TabIndex = 729;
			this.labelEn8.Text = "Estado";
			this.labelEn8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn8.Visible = false;
			// 
			// labelEn7
			// 
			this.labelEn7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn7.Location = new System.Drawing.Point(936, 32);
			this.labelEn7.Name = "labelEn7";
			this.labelEn7.Size = new System.Drawing.Size(120, 16);
			this.labelEn7.TabIndex = 730;
			this.labelEn7.Text = "Estado";
			this.labelEn7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn7.Visible = false;
			// 
			// labelEn43
			// 
			this.labelEn43.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn43.Location = new System.Drawing.Point(328, 472);
			this.labelEn43.Name = "labelEn43";
			this.labelEn43.Size = new System.Drawing.Size(120, 16);
			this.labelEn43.TabIndex = 731;
			this.labelEn43.Text = "Estado";
			this.labelEn43.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn43.Visible = false;
			this.labelEn43.Click += new System.EventHandler(this.label30_Click);
			// 
			// labelEn42
			// 
			this.labelEn42.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn42.Location = new System.Drawing.Point(176, 472);
			this.labelEn42.Name = "labelEn42";
			this.labelEn42.Size = new System.Drawing.Size(120, 16);
			this.labelEn42.TabIndex = 732;
			this.labelEn42.Text = "Estado";
			this.labelEn42.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn42.Visible = false;
			// 
			// labelEn41
			// 
			this.labelEn41.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn41.Location = new System.Drawing.Point(24, 472);
			this.labelEn41.Name = "labelEn41";
			this.labelEn41.Size = new System.Drawing.Size(120, 16);
			this.labelEn41.TabIndex = 733;
			this.labelEn41.Text = "Estado";
			this.labelEn41.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn41.Visible = false;
			this.labelEn41.Click += new System.EventHandler(this.label32_Click);
			// 
			// labelEn40
			// 
			this.labelEn40.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn40.Location = new System.Drawing.Point(1088, 384);
			this.labelEn40.Name = "labelEn40";
			this.labelEn40.Size = new System.Drawing.Size(120, 16);
			this.labelEn40.TabIndex = 734;
			this.labelEn40.Text = "Estado";
			this.labelEn40.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn40.Visible = false;
			// 
			// labelEn39
			// 
			this.labelEn39.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn39.Location = new System.Drawing.Point(936, 384);
			this.labelEn39.Name = "labelEn39";
			this.labelEn39.Size = new System.Drawing.Size(120, 16);
			this.labelEn39.TabIndex = 735;
			this.labelEn39.Text = "Estado";
			this.labelEn39.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn39.Visible = false;
			this.labelEn39.Click += new System.EventHandler(this.label34_Click);
			// 
			// labelEn38
			// 
			this.labelEn38.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn38.Location = new System.Drawing.Point(784, 384);
			this.labelEn38.Name = "labelEn38";
			this.labelEn38.Size = new System.Drawing.Size(120, 16);
			this.labelEn38.TabIndex = 736;
			this.labelEn38.Text = "Estado";
			this.labelEn38.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn38.Visible = false;
			// 
			// labelEn37
			// 
			this.labelEn37.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn37.Location = new System.Drawing.Point(632, 384);
			this.labelEn37.Name = "labelEn37";
			this.labelEn37.Size = new System.Drawing.Size(120, 16);
			this.labelEn37.TabIndex = 737;
			this.labelEn37.Text = "Estado";
			this.labelEn37.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn37.Visible = false;
			// 
			// labelEn36
			// 
			this.labelEn36.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn36.Location = new System.Drawing.Point(480, 384);
			this.labelEn36.Name = "labelEn36";
			this.labelEn36.Size = new System.Drawing.Size(120, 16);
			this.labelEn36.TabIndex = 738;
			this.labelEn36.Text = "Estado";
			this.labelEn36.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn36.Visible = false;
			// 
			// labelEn35
			// 
			this.labelEn35.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn35.Location = new System.Drawing.Point(328, 384);
			this.labelEn35.Name = "labelEn35";
			this.labelEn35.Size = new System.Drawing.Size(120, 16);
			this.labelEn35.TabIndex = 739;
			this.labelEn35.Text = "Estado";
			this.labelEn35.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn35.Visible = false;
			// 
			// labelEn34
			// 
			this.labelEn34.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn34.Location = new System.Drawing.Point(176, 384);
			this.labelEn34.Name = "labelEn34";
			this.labelEn34.Size = new System.Drawing.Size(120, 16);
			this.labelEn34.TabIndex = 740;
			this.labelEn34.Text = "Estado";
			this.labelEn34.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn34.Visible = false;
			// 
			// labelEn33
			// 
			this.labelEn33.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn33.Location = new System.Drawing.Point(24, 384);
			this.labelEn33.Name = "labelEn33";
			this.labelEn33.Size = new System.Drawing.Size(120, 16);
			this.labelEn33.TabIndex = 741;
			this.labelEn33.Text = "Estado";
			this.labelEn33.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn33.Visible = false;
			// 
			// labelEn32
			// 
			this.labelEn32.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn32.Location = new System.Drawing.Point(1088, 296);
			this.labelEn32.Name = "labelEn32";
			this.labelEn32.Size = new System.Drawing.Size(120, 16);
			this.labelEn32.TabIndex = 742;
			this.labelEn32.Text = "Estado";
			this.labelEn32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn32.Visible = false;
			// 
			// labelEn31
			// 
			this.labelEn31.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn31.Location = new System.Drawing.Point(936, 296);
			this.labelEn31.Name = "labelEn31";
			this.labelEn31.Size = new System.Drawing.Size(120, 16);
			this.labelEn31.TabIndex = 743;
			this.labelEn31.Text = "Estado";
			this.labelEn31.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn31.Visible = false;
			this.labelEn31.Click += new System.EventHandler(this.label42_Click);
			// 
			// labelEn30
			// 
			this.labelEn30.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn30.Location = new System.Drawing.Point(784, 296);
			this.labelEn30.Name = "labelEn30";
			this.labelEn30.Size = new System.Drawing.Size(120, 16);
			this.labelEn30.TabIndex = 744;
			this.labelEn30.Text = "Estado";
			this.labelEn30.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn30.Visible = false;
			// 
			// labelEn48
			// 
			this.labelEn48.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn48.Location = new System.Drawing.Point(1088, 472);
			this.labelEn48.Name = "labelEn48";
			this.labelEn48.Size = new System.Drawing.Size(120, 16);
			this.labelEn48.TabIndex = 745;
			this.labelEn48.Text = "Estado";
			this.labelEn48.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn48.Visible = false;
			// 
			// labelEn47
			// 
			this.labelEn47.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn47.Location = new System.Drawing.Point(936, 472);
			this.labelEn47.Name = "labelEn47";
			this.labelEn47.Size = new System.Drawing.Size(120, 16);
			this.labelEn47.TabIndex = 746;
			this.labelEn47.Text = "Estado";
			this.labelEn47.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn47.Visible = false;
			// 
			// labelEn46
			// 
			this.labelEn46.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn46.Location = new System.Drawing.Point(784, 472);
			this.labelEn46.Name = "labelEn46";
			this.labelEn46.Size = new System.Drawing.Size(120, 16);
			this.labelEn46.TabIndex = 747;
			this.labelEn46.Text = "Estado";
			this.labelEn46.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn46.Visible = false;
			// 
			// labelEn45
			// 
			this.labelEn45.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn45.Location = new System.Drawing.Point(632, 472);
			this.labelEn45.Name = "labelEn45";
			this.labelEn45.Size = new System.Drawing.Size(120, 16);
			this.labelEn45.TabIndex = 748;
			this.labelEn45.Text = "Estado";
			this.labelEn45.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn45.Visible = false;
			// 
			// labelEn44
			// 
			this.labelEn44.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelEn44.Location = new System.Drawing.Point(480, 472);
			this.labelEn44.Name = "labelEn44";
			this.labelEn44.Size = new System.Drawing.Size(120, 16);
			this.labelEn44.TabIndex = 749;
			this.labelEn44.Text = "Estado";
			this.labelEn44.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.labelEn44.Visible = false;
			// 
			// frmCuadriculaHabitaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1232, 542);
			this.Controls.Add(this.labelEn44);
			this.Controls.Add(this.labelEn45);
			this.Controls.Add(this.labelEn46);
			this.Controls.Add(this.labelEn47);
			this.Controls.Add(this.labelEn48);
			this.Controls.Add(this.labelEn30);
			this.Controls.Add(this.labelEn31);
			this.Controls.Add(this.labelEn32);
			this.Controls.Add(this.labelEn33);
			this.Controls.Add(this.labelEn34);
			this.Controls.Add(this.labelEn35);
			this.Controls.Add(this.labelEn36);
			this.Controls.Add(this.labelEn37);
			this.Controls.Add(this.labelEn38);
			this.Controls.Add(this.labelEn39);
			this.Controls.Add(this.labelEn40);
			this.Controls.Add(this.labelEn41);
			this.Controls.Add(this.labelEn42);
			this.Controls.Add(this.labelEn43);
			this.Controls.Add(this.labelEn7);
			this.Controls.Add(this.labelEn8);
			this.Controls.Add(this.labelEn9);
			this.Controls.Add(this.labelEn10);
			this.Controls.Add(this.labelEn11);
			this.Controls.Add(this.labelEn12);
			this.Controls.Add(this.labelEn13);
			this.Controls.Add(this.labelEn14);
			this.Controls.Add(this.labelEn15);
			this.Controls.Add(this.labelEn16);
			this.Controls.Add(this.labelEn17);
			this.Controls.Add(this.labelEn18);
			this.Controls.Add(this.labelEn19);
			this.Controls.Add(this.labelEn20);
			this.Controls.Add(this.labelEn21);
			this.Controls.Add(this.labelEn22);
			this.Controls.Add(this.labelEn23);
			this.Controls.Add(this.labelEn24);
			this.Controls.Add(this.labelEn25);
			this.Controls.Add(this.labelEn26);
			this.Controls.Add(this.labelEn27);
			this.Controls.Add(this.labelEn28);
			this.Controls.Add(this.labelEn29);
			this.Controls.Add(this.labelEn6);
			this.Controls.Add(this.labelEn5);
			this.Controls.Add(this.labelEn4);
			this.Controls.Add(this.labelEn3);
			this.Controls.Add(this.labelEn2);
			this.Controls.Add(this.labelEn1);
			this.Controls.Add(this.labH144);
			this.Controls.Add(this.labH143);
			this.Controls.Add(this.labH142);
			this.Controls.Add(this.button48);
			this.Controls.Add(this.labH141);
			this.Controls.Add(this.labH140);
			this.Controls.Add(this.labH139);
			this.Controls.Add(this.button47);
			this.Controls.Add(this.labH138);
			this.Controls.Add(this.labH137);
			this.Controls.Add(this.labH136);
			this.Controls.Add(this.button46);
			this.Controls.Add(this.labH135);
			this.Controls.Add(this.labH134);
			this.Controls.Add(this.labH133);
			this.Controls.Add(this.button45);
			this.Controls.Add(this.labH132);
			this.Controls.Add(this.labH131);
			this.Controls.Add(this.labH130);
			this.Controls.Add(this.button44);
			this.Controls.Add(this.labH129);
			this.Controls.Add(this.labH128);
			this.Controls.Add(this.labH127);
			this.Controls.Add(this.button43);
			this.Controls.Add(this.labH126);
			this.Controls.Add(this.labH125);
			this.Controls.Add(this.labH124);
			this.Controls.Add(this.button42);
			this.Controls.Add(this.labH123);
			this.Controls.Add(this.labH122);
			this.Controls.Add(this.labH121);
			this.Controls.Add(this.button41);
			this.Controls.Add(this.labH120);
			this.Controls.Add(this.labH119);
			this.Controls.Add(this.labH118);
			this.Controls.Add(this.button40);
			this.Controls.Add(this.labH117);
			this.Controls.Add(this.labH116);
			this.Controls.Add(this.labH115);
			this.Controls.Add(this.button39);
			this.Controls.Add(this.labH114);
			this.Controls.Add(this.labH113);
			this.Controls.Add(this.labH112);
			this.Controls.Add(this.button38);
			this.Controls.Add(this.labH111);
			this.Controls.Add(this.labH110);
			this.Controls.Add(this.labH109);
			this.Controls.Add(this.button37);
			this.Controls.Add(this.labH108);
			this.Controls.Add(this.labH107);
			this.Controls.Add(this.labH106);
			this.Controls.Add(this.button36);
			this.Controls.Add(this.labH105);
			this.Controls.Add(this.labH104);
			this.Controls.Add(this.labH103);
			this.Controls.Add(this.button35);
			this.Controls.Add(this.labH102);
			this.Controls.Add(this.labH101);
			this.Controls.Add(this.labH100);
			this.Controls.Add(this.button34);
			this.Controls.Add(this.labH99);
			this.Controls.Add(this.labH98);
			this.Controls.Add(this.labH97);
			this.Controls.Add(this.button33);
			this.Controls.Add(this.labH96);
			this.Controls.Add(this.labH95);
			this.Controls.Add(this.labH94);
			this.Controls.Add(this.button32);
			this.Controls.Add(this.labH93);
			this.Controls.Add(this.labH92);
			this.Controls.Add(this.labH91);
			this.Controls.Add(this.button31);
			this.Controls.Add(this.labH90);
			this.Controls.Add(this.labH89);
			this.Controls.Add(this.labH88);
			this.Controls.Add(this.button30);
			this.Controls.Add(this.labH87);
			this.Controls.Add(this.labH86);
			this.Controls.Add(this.labH85);
			this.Controls.Add(this.button29);
			this.Controls.Add(this.labH84);
			this.Controls.Add(this.labH83);
			this.Controls.Add(this.labH82);
			this.Controls.Add(this.button28);
			this.Controls.Add(this.labH81);
			this.Controls.Add(this.labH80);
			this.Controls.Add(this.labH79);
			this.Controls.Add(this.button27);
			this.Controls.Add(this.labH78);
			this.Controls.Add(this.labH77);
			this.Controls.Add(this.labH76);
			this.Controls.Add(this.button26);
			this.Controls.Add(this.labH75);
			this.Controls.Add(this.labH74);
			this.Controls.Add(this.labH73);
			this.Controls.Add(this.button25);
			this.Controls.Add(this.labH72);
			this.Controls.Add(this.labH71);
			this.Controls.Add(this.labH70);
			this.Controls.Add(this.button24);
			this.Controls.Add(this.labH69);
			this.Controls.Add(this.labH68);
			this.Controls.Add(this.labH67);
			this.Controls.Add(this.button23);
			this.Controls.Add(this.labH66);
			this.Controls.Add(this.labH65);
			this.Controls.Add(this.labH64);
			this.Controls.Add(this.button22);
			this.Controls.Add(this.labH63);
			this.Controls.Add(this.labH62);
			this.Controls.Add(this.labH61);
			this.Controls.Add(this.button21);
			this.Controls.Add(this.labH60);
			this.Controls.Add(this.labH59);
			this.Controls.Add(this.labH58);
			this.Controls.Add(this.button20);
			this.Controls.Add(this.labH57);
			this.Controls.Add(this.labH56);
			this.Controls.Add(this.labH55);
			this.Controls.Add(this.button19);
			this.Controls.Add(this.labH54);
			this.Controls.Add(this.labH53);
			this.Controls.Add(this.labH52);
			this.Controls.Add(this.button18);
			this.Controls.Add(this.labH51);
			this.Controls.Add(this.labH50);
			this.Controls.Add(this.labH49);
			this.Controls.Add(this.button17);
			this.Controls.Add(this.labH48);
			this.Controls.Add(this.labH47);
			this.Controls.Add(this.labH46);
			this.Controls.Add(this.button16);
			this.Controls.Add(this.labH45);
			this.Controls.Add(this.labH44);
			this.Controls.Add(this.labH43);
			this.Controls.Add(this.button15);
			this.Controls.Add(this.labH42);
			this.Controls.Add(this.labH41);
			this.Controls.Add(this.labH40);
			this.Controls.Add(this.button14);
			this.Controls.Add(this.labH39);
			this.Controls.Add(this.labH38);
			this.Controls.Add(this.labH37);
			this.Controls.Add(this.button13);
			this.Controls.Add(this.labH36);
			this.Controls.Add(this.labH35);
			this.Controls.Add(this.labH34);
			this.Controls.Add(this.button12);
			this.Controls.Add(this.labH33);
			this.Controls.Add(this.labH32);
			this.Controls.Add(this.labH31);
			this.Controls.Add(this.button11);
			this.Controls.Add(this.labH30);
			this.Controls.Add(this.labH29);
			this.Controls.Add(this.labH28);
			this.Controls.Add(this.button10);
			this.Controls.Add(this.labH27);
			this.Controls.Add(this.labH26);
			this.Controls.Add(this.labH25);
			this.Controls.Add(this.button9);
			this.Controls.Add(this.labH24);
			this.Controls.Add(this.labH23);
			this.Controls.Add(this.labH22);
			this.Controls.Add(this.button8);
			this.Controls.Add(this.labH21);
			this.Controls.Add(this.labH20);
			this.Controls.Add(this.labH19);
			this.Controls.Add(this.button7);
			this.Controls.Add(this.labH18);
			this.Controls.Add(this.labH17);
			this.Controls.Add(this.labH16);
			this.Controls.Add(this.button6);
			this.Controls.Add(this.labH15);
			this.Controls.Add(this.labH14);
			this.Controls.Add(this.labH13);
			this.Controls.Add(this.button5);
			this.Controls.Add(this.labH12);
			this.Controls.Add(this.labH11);
			this.Controls.Add(this.labH10);
			this.Controls.Add(this.button4);
			this.Controls.Add(this.labH9);
			this.Controls.Add(this.labH8);
			this.Controls.Add(this.labH7);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.labH6);
			this.Controls.Add(this.labH5);
			this.Controls.Add(this.labH4);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.labH3);
			this.Controls.Add(this.labH2);
			this.Controls.Add(this.labH1);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.dtFecha);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCuadriculaHabitaciones";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Estado de Habitaciones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCuadriculaHabitaciones_KeyDown);
			this.Load += new System.EventHandler(this.frmCuadriculaHabitaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		private void frmCuadriculaHabitaciones_Load(object sender, System.EventArgs e)
		{
			//Habilita opciones de Timer
			timer1.Interval = 90000; //Dos minutos
			timer1.Start();
			timer1.Enabled = true;

			int sw = 0;

			string sSQL = string.Format("Select Count(*) From UsuariosBodegas Where idUsuario = {0}", MenuLatinium.IdUsuario);
			int CantEmpUser = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			if(CantEmpUser > 1) 
			{				
				using (VentanaBodegas miBusqda = new VentanaBodegas())
				{
					if (DialogResult.OK == miBusqda.ShowDialog())
					{
						iBodega = (int)miBusqda.sCodBodega;
						nombrdlBodega = miBusqda.nmBodeg;
						idBodega = Funcion.iEscalarSQL(cdsSeteoF, "SELECT idBodega FROM Bodega WHERE Bodega="+iBodega.ToString());
					}
					//else this.Close();
				}
			}
			else 
			{
				iBodega = Funcion.iEscalarSQL(cdsSeteoF, "SELECT Bodega FROM UsuariosBodegas WHERE idUsuario="+MenuLatinium.IdUsuario.ToString());
				idBodega = Funcion.iEscalarSQL(cdsSeteoF, "SELECT idBodega FROM Bodega WHERE Bodega="+iBodega.ToString());
				nombrdlBodega = Funcion.sEscalarSQL(cdsSeteoF, "SELECT Nombre FROM Bodega WHERE Bodega="+iBodega.ToString());
			}		

			bool bCobroAnticipado = Funcion.bEscalarSQL( cdsSeteoF,string.Format("SELECT CobroAnticipado FROM Bodega WHERE Bodega = {0}", iBodega),true);
			menuEnvioCB.Visible = bCobroAnticipado;
			menuIngresoCLI.Visible = bCobroAnticipado;
//			this.timer2.Enabled = true;
//			this.timer2.Start();
			
			this.Text = "Estado de Habitaciones : " + nombrdlBodega;

			int CantHabtcns = Funcion.iEscalarSQL(cdsSeteoF, "SELECT Count(*) as CanT FROM Habitaciones WHERE idSucursal="+idBodega.ToString());

			ColorearBotones(this, CantHabtcns);

			ColorearEqiquetas(this, CantHabtcns);

			EstadoHabitaciones();
		}

		public void EstadoHabitaciones()
		{
			int nmHabtcn = 0;
			DateTime HrInicia = DateTime.Now;
			DateTime HrFinlza = DateTime.Now;
			string hini = "";
			string hfin = "";
			int edoHabit = 0;
			int edoEnct = 0;
			int difmints = 0;
			string hrini = "";
			string hrfin = "";
			decimal dPrecio = 0.00m;
			bool bToqueQueda = false;
			
			string lQry = "Exec Estado_Habitaciones "+idBodega.ToString();
			SqlDataReader drEdoHab = Funcion.rEscalarSQL(cdsSeteoF, lQry, true);
			while(drEdoHab.Read())
			{
				nmHabtcn = drEdoHab.GetInt32(0);
				HrInicia = (DateTime)drEdoHab.GetValue(1);
				hini = drEdoHab.GetString(2);
				HrFinlza = (DateTime)drEdoHab.GetValue(3);
				hfin = drEdoHab.GetString(4);
				edoHabit = drEdoHab.GetInt32(5);
				difmints = drEdoHab.GetInt32(6);
				TiempoReser = drEdoHab.GetInt32(7);
				hrini = drEdoHab.GetString(8);
				hrfin = drEdoHab.GetString(9);
				dPrecio = drEdoHab.GetDecimal(10);
				bToqueQueda = drEdoHab.GetBoolean(11);
				edoEnct = drEdoHab.GetInt32(12);
				
				CambiaColorHabitacion(nmHabtcn,hini,hfin,edoHabit,difmints,TiempoReser,hrini,hrfin,dPrecio,bToqueQueda,edoEnct);
			}
			drEdoHab.Close();
		}

		public void CambiaColorHabitacion(int NurmdHabitacion, string Hinicia, string HFinaliza, int e_hab, int dif_mins, int TR,string hri, string hrf, decimal dPrecio, bool bToqueQueda,int e_enc)
		{
			#region habitaciones estado
			#region h1
			if (NurmdHabitacion == 1)
			{
				if (e_hab == 2)
				{
					if (dif_mins >= 0 && dif_mins <= 10) // PASADO HASTA 10 MINUTOS
					{
						this.button1.BackColor = Color.SteelBlue;
						this.labH1.BackColor = Color.SteelBlue;
						this.labH2.BackColor = Color.SteelBlue;
						this.labH3.BackColor = Color.SteelBlue;
						this.labH1.Text = Hinicia;
						this.labH2.Text = HFinaliza;
						this.labH3.Text = "PASADO:"+dif_mins.ToString()+" Min.";						
					}
					else if (dif_mins >= -10 && dif_mins < 0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{
						this.button1.BackColor = Color.Firebrick;
						this.labH1.BackColor = Color.Firebrick;
						this.labH2.BackColor = Color.Firebrick;
						this.labH3.BackColor = Color.Firebrick;
						this.labH1.Text = Hinicia;
						this.labH2.Text = HFinaliza;
						this.labH3.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button1.BackColor = Color.Orange;
							this.labH1.BackColor = Color.Orange;
							this.labH2.BackColor = Color.Orange;
							this.labH3.BackColor = Color.Orange;
							this.labH1.Text = Hinicia;
							this.labH2.Text = HFinaliza;
							this.labH3.Text = "T. QUEDA";
						}
						else
						{
							this.button1.BackColor = Color.Red;
							this.labH1.BackColor = Color.Red;
							this.labH2.BackColor = Color.Red;
							this.labH3.BackColor = Color.Red;
							this.labH1.Text = Hinicia;
							this.labH2.Text = HFinaliza;
							this.labH3.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{					
					{
						this.button1.BackColor = Color.Yellow;
						this.button1.ForeColor = Color.Black;
						this.labH1.BackColor = Color.Yellow;
						this.labH2.BackColor = Color.Yellow;
						this.labH3.BackColor = Color.Yellow;
						this.labH1.ForeColor = Color.Black;
						this.labH2.ForeColor = Color.Black;
						this.labH3.ForeColor = Color.Black;
						this.labH1.Text = Hinicia;
						this.labH2.Text = HFinaliza;
						this.labH3.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button1.BackColor = Color.Silver;
					this.labH1.BackColor = Color.Silver;
					this.labH2.BackColor = Color.Silver;
					this.labH3.BackColor = Color.Silver;
					this.labH3.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button1.BackColor = Color.SaddleBrown;
						this.labH1.BackColor = Color.SaddleBrown;
						this.labH2.BackColor = Color.SaddleBrown;
						this.labH3.BackColor = Color.SaddleBrown;
						this.labH1.Text = hri;
						this.labH2.Text = hrf;
						this.labH3.Text = "RESERVADO";
					}
					else
					{
						this.button1.BackColor = Color.Green;
						this.labH1.BackColor = Color.Green;
						this.labH2.BackColor = Color.Green;
						this.labH3.BackColor = Color.Green;
						this.labH1.Text = "00:00";
						this.labH2.Text = "00:00";
						this.labH3.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn1.Visible = true;
					this.labelEn1.BackColor = Color.Red;
					this.labelEn1.ForeColor = Color.White;
					this.labelEn1.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn1.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button1.BackColor = Color.MediumPurple;
					this.labH1.BackColor = Color.MediumPurple;
					this.labH2.BackColor = Color.MediumPurple;
					this.labH3.BackColor = Color.MediumPurple;
					this.labH1.Text = "";
					this.labH2.Text = "";
					this.labH3.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH3.Text = "INGRESO CLIENTE";
					
				}
				if (e_hab == 6)
				{
					this.button1.BackColor = Color.DarkSlateBlue;
					this.labH1.BackColor = Color.DarkSlateBlue;
					this.labH2.BackColor = Color.DarkSlateBlue;
					this.labH3.BackColor = Color.DarkSlateBlue;
					this.labH1.Text = "";
					this.labH2.Text = "";
					this.labH3.Font = new Font(this.labH3.Font.FontFamily, 11);
					this.labH3.Text = "ENVÍO COBRO";
					
				}
				if (e_hab == 7)
				{
					this.button1.BackColor = Color.Turquoise;
					this.labH1.BackColor = Color.Turquoise;
					this.labH2.BackColor = Color.Turquoise;
					this.labH3.BackColor = Color.Turquoise;
					this.button1.ForeColor = Color.Black;
					this.labH1.ForeColor = Color.Black;
					this.labH2.ForeColor = Color.Black;
					this.labH3.ForeColor = Color.Black;
					this.labH1.Text = "";
					this.labH2.Text = "";
					this.labH3.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH3.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button1.BackColor = Color.Navy;
					this.labH1.BackColor = Color.Navy;
					this.labH2.BackColor = Color.Navy;
					this.labH3.BackColor = Color.Navy;
					this.button1.ForeColor = Color.White;
					this.labH1.ForeColor = Color.White;
					this.labH2.ForeColor = Color.White;
					this.labH3.ForeColor = Color.White;
					this.labH1.Text = "";
					this.labH2.Text = "";
					this.labH3.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH3.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button1.BackColor = Color.OrangeRed;
					this.labH1.BackColor = Color.OrangeRed;
					this.labH2.BackColor = Color.OrangeRed;
					this.labH3.BackColor = Color.OrangeRed;
					this.button1.ForeColor = Color.White;
					this.labH1.ForeColor = Color.White;
					this.labH2.ForeColor = Color.White;
					this.labH3.ForeColor = Color.White;
					this.labH1.Text = "";
					this.labH2.Text = "";
					this.labH3.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH3.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button1.BackColor = Color.Orange;
					this.labH1.BackColor = Color.Orange;
					this.labH2.BackColor = Color.Orange;
					this.labH3.BackColor = Color.Orange;
					this.button1.ForeColor = Color.White;
					this.labH1.ForeColor = Color.White;
					this.labH2.ForeColor = Color.White;
					this.labH3.ForeColor = Color.White;
					this.labH1.Text = "";
					this.labH2.Text = "";
					this.labH3.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH3.Text = "EN PROCESO DE SALIDA";
				}
				//					this.lblPrecioH1.Text = dPrecio.ToString("##0.00");
			}
			#endregion h1
			#region h2
				if (NurmdHabitacion==2)
				{
					if (e_hab == 2)
					{
						if (dif_mins >= 0 && dif_mins <= 10) // PASADO HASTA 10 MINUTOS
						{
							this.button2.BackColor = Color.SteelBlue;
							this.labH4.BackColor = Color.SteelBlue;
							this.labH5.BackColor = Color.SteelBlue;
							this.labH6.BackColor = Color.SteelBlue;
							this.labH4.Text = Hinicia;
							this.labH5.Text = HFinaliza;
							this.labH6.Text = "PASADO:"+dif_mins.ToString()+" Min.";						
						}
						else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
						{							
							this.button2.BackColor = Color.Firebrick;
							this.labH4.BackColor = Color.Firebrick;
							this.labH5.BackColor = Color.Firebrick;
							this.labH6.BackColor = Color.Firebrick;
							this.labH4.Text = Hinicia;
							this.labH5.Text = HFinaliza;
							this.labH6.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
						}
						else // OCUPADA
						{
							if(bToqueQueda)
							{
								this.button2.BackColor = Color.Orange;
								this.labH4.BackColor = Color.Orange;
								this.labH5.BackColor = Color.Orange;
								this.labH6.BackColor = Color.Orange;
								this.labH4.Text = Hinicia;
								this.labH5.Text = HFinaliza;
								this.labH6.Text = "T. QUEDA";
							}
							else
							{
								this.button2.BackColor = Color.Red;
								this.labH4.BackColor = Color.Red;
								this.labH5.BackColor = Color.Red;
								this.labH6.BackColor = Color.Red;
								this.labH4.Text = Hinicia;
								this.labH5.Text = HFinaliza;
								this.labH6.Text = "OCUPADO";
							}
						}
					}
					if (e_hab==3)
					{
						
						{
							this.button2.BackColor = Color.Yellow;
							this.button2.ForeColor = Color.Black;
							this.labH4.BackColor = Color.Yellow;
							this.labH5.BackColor = Color.Yellow;
							this.labH6.BackColor = Color.Yellow;
							this.labH4.ForeColor = Color.Black;
							this.labH5.ForeColor = Color.Black;
							this.labH6.ForeColor = Color.Black;
							this.labH4.Text = Hinicia;
							this.labH5.Text = HFinaliza;
							this.labH6.Text = "LIMPIEZA";
						}
					}
					if (e_hab==4)
					{
						this.button2.BackColor = Color.Silver;
						this.labH4.BackColor = Color.Silver;
						this.labH5.BackColor = Color.Silver;
						this.labH6.BackColor = Color.Silver;
						this.labH6.Text = "MANTENIMIENTO";
					}
					if (e_hab==1)
					{
						if (TR>0)
						{
							this.button2.BackColor = Color.SaddleBrown;
							this.labH4.BackColor = Color.SaddleBrown;
							this.labH5.BackColor = Color.SaddleBrown;
							this.labH6.BackColor = Color.SaddleBrown;
							this.labH4.Text = hri;
							this.labH5.Text = hrf;
							this.labH6.Text = "RESERVADO";
						}
						else
						{
							this.button2.BackColor = Color.Green;
							this.labH4.BackColor = Color.Green;
							this.labH5.BackColor = Color.Green;
							this.labH6.BackColor = Color.Green;
							this.labH4.Text = "00:00";
							this.labH5.Text = "00:00";
							this.labH6.Text = "$ "+dPrecio.ToString()+" LIBRE";
						}
					}
					if (e_enc == 1)
					{
						this.labelEn2.Visible = true;
						this.labelEn2.BackColor = Color.Red;
						this.labelEn2.ForeColor = Color.White;
						this.labelEn2.Text = "ENCUESTA";
					}
					else
					{
						this.labelEn2.Visible = false;
					}
					if (e_hab == 5)
					{
						this.button2.BackColor = Color.MediumPurple;
						this.labH4.BackColor = Color.MediumPurple;
						this.labH5.BackColor = Color.MediumPurple;
						this.labH6.BackColor = Color.MediumPurple;
						this.labH4.Text = "";
						this.labH5.Text = "";
						this.labH6.Font = new Font(this.labH6.Font.FontFamily, 10);
						this.labH6.Text = "INGRESO CLIENTE";
					}

					if (e_hab == 6)
					{
						this.button2.BackColor = Color.DarkSlateBlue;
						this.labH4.BackColor = Color.DarkSlateBlue;
						this.labH5.BackColor = Color.DarkSlateBlue;
						this.labH6.BackColor = Color.DarkSlateBlue;
						this.labH4.Text = "";
						this.labH5.Text = "";
						this.labH6.Font = new Font(this.labH6.Font.FontFamily, 11);
						this.labH6.Text = "ENVÍO COBRO";
					}
					if (e_hab == 7)
					{
						this.button2.BackColor = Color.Turquoise;
						this.labH4.BackColor = Color.Turquoise;
						this.labH5.BackColor = Color.Turquoise;
						this.labH6.BackColor = Color.Turquoise;
						this.button2.ForeColor = Color.Black;
						this.labH4.ForeColor = Color.Black;
						this.labH5.ForeColor = Color.Black;
						this.labH6.ForeColor = Color.Black;
						this.labH4.Text = "";
						this.labH5.Text = "";
						this.labH6.Font = new Font(this.labH6.Font.FontFamily, 10);
						this.labH6.Text = "PROCESO LIMPIEZA";	
					}
					if (e_hab == 8)
					{
						this.button2.BackColor = Color.Navy;
						this.labH4.BackColor = Color.Navy;
						this.labH5.BackColor = Color.Navy;
						this.labH6.BackColor = Color.Navy;
						this.button2.ForeColor = Color.White;
						this.labH4.ForeColor = Color.White;
						this.labH5.ForeColor = Color.White;
						this.labH6.ForeColor = Color.White;
						this.labH4.Text = "";
						this.labH5.Text = "";
						this.labH6.Font = new Font(this.labH6.Font.FontFamily, 10);
						this.labH6.Text = "APROBAR LIMPIEZA";				
					}
					if (e_hab == 9)
					{
						this.button2.BackColor = Color.OrangeRed;
						this.labH4.BackColor = Color.OrangeRed;
						this.labH5.BackColor = Color.OrangeRed;
						this.labH6.BackColor = Color.OrangeRed;
						this.button2.ForeColor = Color.White;
						this.labH4.ForeColor = Color.White;
						this.labH5.ForeColor = Color.White;
						this.labH6.ForeColor = Color.White;
						this.labH4.Text = "";
						this.labH5.Text = "";
						this.labH6.Font = new Font(this.labH6.Font.FontFamily, 10);
						this.labH6.Text = "CORREGIR LIMPIEZA";
					}
					if (e_hab == 10)
					{
						this.button2.BackColor = Color.Orange;
						this.labH4.BackColor = Color.Orange;
						this.labH5.BackColor = Color.Orange;
						this.labH6.BackColor = Color.Orange;
						this.button2.ForeColor = Color.White;
						this.labH4.ForeColor = Color.White;
						this.labH5.ForeColor = Color.White;
						this.labH6.ForeColor = Color.White;
						this.labH4.Text = "";
						this.labH5.Text = "";
						this.labH6.Font = new Font(this.labH3.Font.FontFamily, 10);
						this.labH6.Text = "EN PROCESO DE SALIDA";
					}
				}
			#endregion h2
			#region h3
			if (NurmdHabitacion == 3)
			{
				if (e_hab == 2)
				{
					if (dif_mins >= 0 && dif_mins <= 10) // PASADO HASTA 10 MINUTOS
					{
						this.button3.BackColor = Color.SteelBlue;
						this.labH7.BackColor = Color.SteelBlue;
						this.labH8.BackColor = Color.SteelBlue;
						this.labH9.BackColor = Color.SteelBlue;
						this.labH7.Text = Hinicia;
						this.labH8.Text = HFinaliza;
						this.labH9.Text = "PASADO:" + dif_mins.ToString() + " Min.";
					}
					else if (dif_mins >= -10 && dif_mins < 0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{
						this.button3.BackColor = Color.Firebrick;
						this.labH7.BackColor = Color.Firebrick;
						this.labH8.BackColor = Color.Firebrick;
						this.labH9.BackColor = Color.Firebrick;
						this.labH7.Text = Hinicia;
						this.labH8.Text = HFinaliza;
						this.labH9.Text = "FALTAN:" + dif_mins.ToString() + " Min.";
					}
					else // OCUPADA
					{
						if (bToqueQueda)
						{
							this.button3.BackColor = Color.Orange;
							this.labH7.BackColor = Color.Orange;
							this.labH8.BackColor = Color.Orange;
							this.labH9.BackColor = Color.Orange;
							this.labH7.Text = Hinicia;
							this.labH8.Text = HFinaliza;
							this.labH9.Text = "T. QUEDA";
						}
						else
						{
							this.button3.BackColor = Color.Red;
							this.labH7.BackColor = Color.Red;
							this.labH8.BackColor = Color.Red;
							this.labH9.BackColor = Color.Red;
							this.labH7.Text = Hinicia;
							this.labH8.Text = HFinaliza;
							this.labH9.Text = "OCUPADO";
						}
					}
				}
				if (e_hab == 3)
				{
					if (dif_mins >= 5)
					{
						this.button3.BackColor = Color.Green;
						this.labH7.BackColor = Color.Green;
						this.labH8.BackColor = Color.Green;
						this.labH9.BackColor = Color.Green;
						this.labH7.Text = "00:00";
						this.labH8.Text = "00:00";
						this.labH9.Text = "$ " + dPrecio.ToString() + " LIBRE";
					}
					else
					{
						this.button3.BackColor = Color.Yellow;
						this.button3.ForeColor = Color.Black;
						this.labH7.BackColor = Color.Yellow;
						this.labH8.BackColor = Color.Yellow;
						this.labH9.BackColor = Color.Yellow;
						this.labH7.ForeColor = Color.Black;
						this.labH8.ForeColor = Color.Black;
						this.labH9.ForeColor = Color.Black;
						this.labH7.Text = Hinicia;
						this.labH8.Text = HFinaliza;
						this.labH9.Text = "LIMPIEZA";
					}
				}
				if (e_hab == 4)
				{
					this.button3.BackColor = Color.Silver;
					this.labH7.BackColor = Color.Silver;
					this.labH8.BackColor = Color.Silver;
					this.labH9.BackColor = Color.Silver;
					this.labH9.Text = "MANTENIMIENTO";
				}
				if (e_hab == 1)
				{
					if (TR > 0)
					{
						this.button3.BackColor = Color.SaddleBrown;
						this.labH7.BackColor = Color.SaddleBrown;
						this.labH8.BackColor = Color.SaddleBrown;
						this.labH9.BackColor = Color.SaddleBrown;
						this.labH7.Text = hri;
						this.labH8.Text = hrf;
						this.labH9.Text = "RESERVADO";
					}
					else
					{
						this.button3.BackColor = Color.Green;
						this.labH7.BackColor = Color.Green;
						this.labH8.BackColor = Color.Green;
						this.labH9.BackColor = Color.Green;
						this.labH7.Text = "00:00";
						this.labH8.Text = "00:00";
						this.labH9.Text = "$ " + dPrecio.ToString() + " LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn3.Visible = true;
					this.labelEn3.BackColor = Color.Red;
					this.labelEn3.ForeColor = Color.White;
					this.labelEn3.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn3.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button3.BackColor = Color.MediumPurple;
					this.labH7.BackColor = Color.MediumPurple;
					this.labH8.BackColor = Color.MediumPurple;
					this.labH9.BackColor = Color.MediumPurple;
					this.labH7.Text = "";
					this.labH8.Text = "";
					this.labH9.Font = new Font(this.labH9.Font.FontFamily, 10);
					this.labH9.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button3.BackColor = Color.DarkSlateBlue;
					this.labH7.BackColor = Color.DarkSlateBlue;
					this.labH8.BackColor = Color.DarkSlateBlue;
					this.labH9.BackColor = Color.DarkSlateBlue;
					this.labH7.Text = "";
					this.labH8.Text = "";
					this.labH9.Font = new Font(this.labH9.Font.FontFamily, 11);
					this.labH9.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button3.BackColor = Color.Turquoise;
					this.labH7.BackColor = Color.Turquoise;
					this.labH8.BackColor = Color.Turquoise;
					this.labH9.BackColor = Color.Turquoise;
					this.button3.ForeColor = Color.Black;
					this.labH7.ForeColor = Color.Black;
					this.labH8.ForeColor = Color.Black;
					this.labH9.ForeColor = Color.Black;
					this.labH7.Text = "";
					this.labH8.Text = "";
					this.labH9.Font = new Font(this.labH9.Font.FontFamily, 10);
					this.labH9.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button3.BackColor = Color.Navy;
					this.labH7.BackColor = Color.Navy;
					this.labH8.BackColor = Color.Navy;
					this.labH9.BackColor = Color.Navy;
					this.button3.ForeColor = Color.White;
					this.labH7.ForeColor = Color.White;
					this.labH8.ForeColor = Color.White;
					this.labH9.ForeColor = Color.White;
					this.labH7.Text = "";
					this.labH8.Text = "";
					this.labH9.Font = new Font(this.labH9.Font.FontFamily, 10);
					this.labH9.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button3.BackColor = Color.OrangeRed;
					this.labH7.BackColor = Color.OrangeRed;
					this.labH8.BackColor = Color.OrangeRed;
					this.labH9.BackColor = Color.OrangeRed;
					this.button3.ForeColor = Color.White;
					this.labH7.ForeColor = Color.White;
					this.labH8.ForeColor = Color.White;
					this.labH9.ForeColor = Color.White;
					this.labH7.Text = "";
					this.labH8.Text = "";
					this.labH9.Font = new Font(this.labH9.Font.FontFamily, 10);
					this.labH9.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button3.BackColor = Color.Orange;
					this.labH7.BackColor = Color.Orange;
					this.labH8.BackColor = Color.Orange;
					this.labH9.BackColor = Color.Orange;
					this.button3.ForeColor = Color.White;
					this.labH7.ForeColor = Color.White;
					this.labH8.ForeColor = Color.White;
					this.labH9.ForeColor = Color.White;
					this.labH7.Text = "";
					this.labH8.Text = "";
					this.labH9.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH9.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h3
			#region h4
			if (NurmdHabitacion==4)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button4.BackColor = Color.SteelBlue;
						this.labH10.BackColor = Color.SteelBlue;
						this.labH11.BackColor = Color.SteelBlue;
						this.labH12.BackColor = Color.SteelBlue;
						this.labH10.Text = Hinicia;
						this.labH11.Text = HFinaliza;
						this.labH12.Text = "PASADO:"+dif_mins.ToString()+" Min.";					
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button4.BackColor = Color.Firebrick;
						this.labH10.BackColor = Color.Firebrick;
						this.labH11.BackColor = Color.Firebrick;
						this.labH12.BackColor = Color.Firebrick;
						this.labH10.Text = Hinicia;
						this.labH11.Text = HFinaliza;
						this.labH12.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button4.BackColor = Color.Orange;
							this.labH10.BackColor = Color.Orange;
							this.labH11.BackColor = Color.Orange;
							this.labH12.BackColor = Color.Orange;
							this.labH10.Text = Hinicia;
							this.labH11.Text = HFinaliza;
							this.labH12.Text = "T. QUEDA";
						}
						else
						{
							this.button4.BackColor = Color.Red;
							this.labH10.BackColor = Color.Red;
							this.labH11.BackColor = Color.Red;
							this.labH12.BackColor = Color.Red;
							this.labH10.Text = Hinicia;
							this.labH11.Text = HFinaliza;
							this.labH12.Text = "OCUPADO";
						}
					}	
				}
				if (e_hab==3)
				{
					
						this.button4.BackColor = Color.Yellow;
						this.button4.ForeColor = Color.Black;
						this.labH10.BackColor = Color.Yellow;
						this.labH11.BackColor = Color.Yellow;
						this.labH12.BackColor = Color.Yellow;
						this.labH10.ForeColor = Color.Black;
						this.labH11.ForeColor = Color.Black;
						this.labH12.ForeColor = Color.Black;
						this.labH10.Text = Hinicia;
						this.labH11.Text = HFinaliza;
						this.labH12.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button4.BackColor = Color.Silver;
					this.labH10.BackColor = Color.Silver;
					this.labH11.BackColor = Color.Silver;
					this.labH12.BackColor = Color.Silver;
					this.labH12.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button4.BackColor = Color.SaddleBrown;
						this.labH10.BackColor = Color.SaddleBrown;
						this.labH11.BackColor = Color.SaddleBrown;
						this.labH12.BackColor = Color.SaddleBrown;
						this.labH10.Text = hri;
						this.labH11.Text = hrf;
						this.labH12.Text = "RESERVADO";
					}
					else
					{
						this.button4.BackColor = Color.Green;
						this.labH10.BackColor = Color.Green;
						this.labH11.BackColor = Color.Green;
						this.labH12.BackColor = Color.Green;
						this.labH10.Text = "00:00";
						this.labH11.Text = "00:00";
						this.labH12.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn4.Visible = true;
					this.labelEn4.BackColor = Color.Red;
					this.labelEn4.ForeColor = Color.White;
					this.labelEn4.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn4.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button4.BackColor = Color.MediumPurple;
					this.labH10.BackColor = Color.MediumPurple;
					this.labH11.BackColor = Color.MediumPurple;
					this.labH12.BackColor = Color.MediumPurple;
					this.labH10.Text = "";
					this.labH11.Text = "";
					this.labH12.Font = new Font(this.labH12.Font.FontFamily, 10);
					this.labH12.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button4.BackColor = Color.DarkSlateBlue;
					this.labH10.BackColor = Color.DarkSlateBlue;
					this.labH11.BackColor = Color.DarkSlateBlue;
					this.labH12.BackColor = Color.DarkSlateBlue;
					this.labH10.Text = "";
					this.labH11.Text = "";
					this.labH12.Font = new Font(this.labH12.Font.FontFamily, 11);
					this.labH12.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button4.BackColor = Color.Turquoise;
					this.labH10.BackColor = Color.Turquoise;
					this.labH11.BackColor = Color.Turquoise;
					this.labH12.BackColor = Color.Turquoise;
					this.button4.ForeColor = Color.Black;
					this.labH10.ForeColor = Color.Black;
					this.labH11.ForeColor = Color.Black;
					this.labH12.ForeColor = Color.Black;
					this.labH10.Text = "";
					this.labH11.Text = "";
					this.labH12.Font = new Font(this.labH12.Font.FontFamily, 10);
					this.labH12.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button4.BackColor = Color.Navy;
					this.labH10.BackColor = Color.Navy;
					this.labH11.BackColor = Color.Navy;
					this.labH12.BackColor = Color.Navy;
					this.button4.ForeColor = Color.White;
					this.labH10.ForeColor = Color.White;
					this.labH11.ForeColor = Color.White;
					this.labH12.ForeColor = Color.White;
					this.labH10.Text = "";
					this.labH11.Text = "";
					this.labH12.Font = new Font(this.labH12.Font.FontFamily, 10);
					this.labH12.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button4.BackColor = Color.OrangeRed;
					this.labH10.BackColor = Color.OrangeRed;
					this.labH11.BackColor = Color.OrangeRed;
					this.labH12.BackColor = Color.OrangeRed;
					this.button4.ForeColor = Color.White;
					this.labH10.ForeColor = Color.White;
					this.labH11.ForeColor = Color.White;
					this.labH12.ForeColor = Color.White;
					this.labH10.Text = "";
					this.labH11.Text = "";
					this.labH12.Font = new Font(this.labH12.Font.FontFamily, 10);
					this.labH12.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button4.BackColor = Color.Orange;
					this.labH10.BackColor = Color.Orange;
					this.labH11.BackColor = Color.Orange;
					this.labH12.BackColor = Color.Orange;
					this.button4.ForeColor = Color.White;
					this.labH10.ForeColor = Color.White;
					this.labH11.ForeColor = Color.White;
					this.labH12.ForeColor = Color.White;
					this.labH10.Text = "";
					this.labH11.Text = "";
					this.labH12.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH12.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h4
			#region h5
			if (NurmdHabitacion==5)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button5.BackColor = Color.SteelBlue;
						this.labH13.BackColor = Color.SteelBlue;
						this.labH14.BackColor = Color.SteelBlue;
						this.labH15.BackColor = Color.SteelBlue;
						this.labH13.Text = Hinicia;
						this.labH14.Text = HFinaliza;
						this.labH15.Text = "PASADO:"+dif_mins.ToString()+" Min.";				
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button5.BackColor = Color.Firebrick;
						this.labH13.BackColor = Color.Firebrick;
						this.labH14.BackColor = Color.Firebrick;
						this.labH15.BackColor = Color.Firebrick;
						this.labH13.Text = Hinicia;
						this.labH14.Text = HFinaliza;
						this.labH15.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button5.BackColor = Color.Orange;
							this.labH13.BackColor = Color.Orange;
							this.labH14.BackColor = Color.Orange;
							this.labH15.BackColor = Color.Orange;
							this.labH13.Text = Hinicia;
							this.labH14.Text = HFinaliza;
							this.labH15.Text = "T. QUEDA";
						}
						else
						{
							this.button5.BackColor = Color.Red;
							this.labH13.BackColor = Color.Red;
							this.labH14.BackColor = Color.Red;
							this.labH15.BackColor = Color.Red;
							this.labH13.Text = Hinicia;
							this.labH14.Text = HFinaliza;
							this.labH15.Text = "OCUPADO";
						}
					}	
				}
				if (e_hab==3)
				{

						this.button5.BackColor = Color.Yellow;
						this.button5.ForeColor = Color.Black;
						this.labH13.BackColor = Color.Yellow;
						this.labH14.BackColor = Color.Yellow;
						this.labH15.BackColor = Color.Yellow;
						this.labH13.ForeColor = Color.Black;
						this.labH14.ForeColor = Color.Black;
						this.labH15.ForeColor = Color.Black;
						this.labH13.Text = Hinicia;
						this.labH14.Text = HFinaliza;
						this.labH15.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button5.BackColor = Color.Silver;
					this.labH13.BackColor = Color.Silver;
					this.labH14.BackColor = Color.Silver;
					this.labH15.BackColor = Color.Silver;
					this.labH15.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button5.BackColor = Color.SaddleBrown;
						this.labH13.BackColor = Color.SaddleBrown;
						this.labH14.BackColor = Color.SaddleBrown;
						this.labH15.BackColor = Color.SaddleBrown;
						this.labH13.Text = hri;
						this.labH14.Text = hrf;
						this.labH15.Text = "RESERVADO";
					}
					else
					{
						this.button5.BackColor = Color.Green;
						this.labH13.BackColor = Color.Green;
						this.labH14.BackColor = Color.Green;
						this.labH15.BackColor = Color.Green;
						this.labH13.Text = "00:00";
						this.labH14.Text = "00:00";
						this.labH15.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn5.Visible = true;
					this.labelEn5.BackColor = Color.Red;
					this.labelEn5.ForeColor = Color.White;
					this.labelEn5.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn5.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button5.BackColor = Color.MediumPurple;
					this.labH13.BackColor = Color.MediumPurple;
					this.labH14.BackColor = Color.MediumPurple;
					this.labH15.BackColor = Color.MediumPurple;
					this.labH13.Text = "";
					this.labH14.Text = "";
					this.labH15.Font = new Font(this.labH15.Font.FontFamily, 10);
					this.labH15.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button5.BackColor = Color.DarkSlateBlue;
					this.labH13.BackColor = Color.DarkSlateBlue;
					this.labH14.BackColor = Color.DarkSlateBlue;
					this.labH15.BackColor = Color.DarkSlateBlue;
					this.labH13.Text = "";
					this.labH14.Text = "";
					this.labH15.Font = new Font(this.labH15.Font.FontFamily, 11);
					this.labH15.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button5.BackColor = Color.Turquoise;
					this.labH13.BackColor = Color.Turquoise;
					this.labH14.BackColor = Color.Turquoise;
					this.labH15.BackColor = Color.Turquoise;
					this.button5.ForeColor = Color.Black;
					this.labH13.ForeColor = Color.Black;
					this.labH14.ForeColor = Color.Black;
					this.labH15.ForeColor = Color.Black;
					this.labH13.Text = "";
					this.labH14.Text = "";
					this.labH15.Font = new Font(this.labH15.Font.FontFamily, 10);
					this.labH15.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button5.BackColor = Color.Navy;
					this.labH13.BackColor = Color.Navy;
					this.labH14.BackColor = Color.Navy;
					this.labH15.BackColor = Color.Navy;
					this.button5.ForeColor = Color.White;
					this.labH13.ForeColor = Color.White;
					this.labH14.ForeColor = Color.White;
					this.labH15.ForeColor = Color.White;
					this.labH13.Text = "";
					this.labH14.Text = "";
					this.labH15.Font = new Font(this.labH15.Font.FontFamily, 10);
					this.labH15.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button5.BackColor = Color.OrangeRed;
					this.labH13.BackColor = Color.OrangeRed;
					this.labH14.BackColor = Color.OrangeRed;
					this.labH15.BackColor = Color.OrangeRed;
					this.button5.ForeColor = Color.White;
					this.labH13.ForeColor = Color.White;
					this.labH14.ForeColor = Color.White;
					this.labH15.ForeColor = Color.White;
					this.labH13.Text = "";
					this.labH14.Text = "";
					this.labH15.Font = new Font(this.labH15.Font.FontFamily, 10);
					this.labH15.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button5.BackColor = Color.Orange;
					this.labH13.BackColor = Color.Orange;
					this.labH14.BackColor = Color.Orange;
					this.labH15.BackColor = Color.Orange;
					this.button5.ForeColor = Color.White;
					this.labH13.ForeColor = Color.White;
					this.labH14.ForeColor = Color.White;
					this.labH15.ForeColor = Color.White;
					this.labH13.Text = "";
					this.labH14.Text = "";
					this.labH15.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH15.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h5
			#region h6
			if (NurmdHabitacion==6)
				{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button6.BackColor = Color.SteelBlue;
						this.labH16.BackColor = Color.SteelBlue;
						this.labH17.BackColor = Color.SteelBlue;
						this.labH18.BackColor = Color.SteelBlue;
						this.labH16.Text = Hinicia;
						this.labH17.Text = HFinaliza;
						this.labH18.Text = "PASADO:"+dif_mins.ToString()+" Min.";			
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button6.BackColor = Color.Firebrick;
						this.labH16.BackColor = Color.Firebrick;
						this.labH17.BackColor = Color.Firebrick;
						this.labH18.BackColor = Color.Firebrick;
						this.labH16.Text = Hinicia;
						this.labH17.Text = HFinaliza;
						this.labH18.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button6.BackColor = Color.Orange;
							this.labH16.BackColor = Color.Orange;
							this.labH17.BackColor = Color.Orange;
							this.labH18.BackColor = Color.Orange;
							this.labH16.Text = Hinicia;
							this.labH17.Text = HFinaliza;
							this.labH18.Text = "T. QUEDA";
						}
						else
						{
							this.button6.BackColor = Color.Red;
							this.labH16.BackColor = Color.Red;
							this.labH17.BackColor = Color.Red;
							this.labH18.BackColor = Color.Red;
							this.labH16.Text = Hinicia;
							this.labH17.Text = HFinaliza;
							this.labH18.Text = "OCUPADO";
						}
					}	
				}
				if (e_hab==3)
				{

						this.button6.BackColor = Color.Yellow;
						this.button6.ForeColor = Color.Black;
						this.labH16.BackColor = Color.Yellow;
						this.labH17.BackColor = Color.Yellow;
						this.labH18.BackColor = Color.Yellow;
						this.labH16.ForeColor = Color.Black;
						this.labH17.ForeColor = Color.Black;
						this.labH18.ForeColor = Color.Black;
						this.labH16.Text = Hinicia;
						this.labH17.Text = HFinaliza;
						this.labH18.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button6.BackColor = Color.Silver;
					this.labH16.BackColor = Color.Silver;
					this.labH17.BackColor = Color.Silver;
					this.labH18.BackColor = Color.Silver;
					this.labH18.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button6.BackColor = Color.SaddleBrown;
						this.labH16.BackColor = Color.SaddleBrown;
						this.labH17.BackColor = Color.SaddleBrown;
						this.labH18.BackColor = Color.SaddleBrown;
						this.labH16.Text = hri;
						this.labH17.Text = hrf;
						this.labH18.Text = "RESERVADO";
					}
					else
					{
						this.button6.BackColor = Color.Green;
						this.labH16.BackColor = Color.Green;
						this.labH17.BackColor = Color.Green;
						this.labH18.BackColor = Color.Green;
						this.labH16.Text = "00:00";
						this.labH17.Text = "00:00";
						this.labH18.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn6.Visible = true;
					this.labelEn6.BackColor = Color.Red;
					this.labelEn6.ForeColor = Color.White;
					this.labelEn6.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn6.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button6.BackColor = Color.MediumPurple;
					this.labH16.BackColor = Color.MediumPurple;
					this.labH17.BackColor = Color.MediumPurple;
					this.labH18.BackColor = Color.MediumPurple;
					this.labH16.Text = "";
					this.labH17.Text = "";
					this.labH18.Font = new Font(this.labH18.Font.FontFamily, 10);
					this.labH18.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button6.BackColor = Color.DarkSlateBlue;
					this.labH16.BackColor = Color.DarkSlateBlue;
					this.labH17.BackColor = Color.DarkSlateBlue;
					this.labH18.BackColor = Color.DarkSlateBlue;
					this.labH16.Text = "";
					this.labH17.Text = "";
					this.labH18.Font = new Font(this.labH18.Font.FontFamily, 11);
					this.labH18.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button6.BackColor = Color.Turquoise;
					this.labH16.BackColor = Color.Turquoise;
					this.labH17.BackColor = Color.Turquoise;
					this.labH18.BackColor = Color.Turquoise;
					this.button6.ForeColor = Color.Black;
					this.labH16.ForeColor = Color.Black;
					this.labH17.ForeColor = Color.Black;
					this.labH18.ForeColor = Color.Black;
					this.labH16.Text = "";
					this.labH17.Text = "";
					this.labH18.Font = new Font(this.labH18.Font.FontFamily, 10);
					this.labH18.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button6.BackColor = Color.Navy;
					this.labH16.BackColor = Color.Navy;
					this.labH17.BackColor = Color.Navy;
					this.labH18.BackColor = Color.Navy;
					this.button6.ForeColor = Color.White;
					this.labH16.ForeColor = Color.White;
					this.labH17.ForeColor = Color.White;
					this.labH18.ForeColor = Color.White;
					this.labH16.Text = "";
					this.labH17.Text = "";
					this.labH18.Font = new Font(this.labH18.Font.FontFamily, 10);
					this.labH18.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button6.BackColor = Color.OrangeRed;
					this.labH16.BackColor = Color.OrangeRed;
					this.labH17.BackColor = Color.OrangeRed;
					this.labH18.BackColor = Color.OrangeRed;
					this.button6.ForeColor = Color.White;
					this.labH16.ForeColor = Color.White;
					this.labH17.ForeColor = Color.White;
					this.labH18.ForeColor = Color.White;
					this.labH16.Text = "";
					this.labH17.Text = "";
					this.labH18.Font = new Font(this.labH18.Font.FontFamily, 10);
					this.labH18.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button6.BackColor = Color.Orange;
					this.labH16.BackColor = Color.Orange;
					this.labH17.BackColor = Color.Orange;
					this.labH18.BackColor = Color.Orange;
					this.button6.ForeColor = Color.White;
					this.labH16.ForeColor = Color.White;
					this.labH17.ForeColor = Color.White;
					this.labH18.ForeColor = Color.White;
					this.labH16.Text = "";
					this.labH17.Text = "";
					this.labH18.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH18.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h6
			#region h7
			if (NurmdHabitacion==7)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button7.BackColor = Color.SteelBlue;
						this.labH19.BackColor = Color.SteelBlue;
						this.labH20.BackColor = Color.SteelBlue;
						this.labH21.BackColor = Color.SteelBlue;
						this.labH19.Text = Hinicia;
						this.labH20.Text = HFinaliza;
						this.labH21.Text = "PASADO:"+dif_mins.ToString()+" Min.";		
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button7.BackColor = Color.Firebrick;
						this.labH19.BackColor = Color.Firebrick;
						this.labH20.BackColor = Color.Firebrick;
						this.labH21.BackColor = Color.Firebrick;
						this.labH19.Text = Hinicia;
						this.labH20.Text = HFinaliza;
						this.labH21.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button7.BackColor = Color.Orange;
							this.labH19.BackColor = Color.Orange;
							this.labH20.BackColor = Color.Orange;
							this.labH21.BackColor = Color.Orange;
							this.labH19.Text = Hinicia;
							this.labH20.Text = HFinaliza;
							this.labH21.Text = "T. QUEDA";
						}
						else
						{
							this.button7.BackColor = Color.Red;
							this.labH19.BackColor = Color.Red;
							this.labH20.BackColor = Color.Red;
							this.labH21.BackColor = Color.Red;
							this.labH19.Text = Hinicia;
							this.labH20.Text = HFinaliza;
							this.labH21.Text = "OCUPADO";
						}
					}	
				}
				if (e_hab==3)
				{

						this.button7.BackColor = Color.Yellow;
						this.button7.ForeColor = Color.Black;
						this.labH19.BackColor = Color.Yellow;
						this.labH20.BackColor = Color.Yellow;
						this.labH21.BackColor = Color.Yellow;
						this.labH20.ForeColor = Color.Black;
						this.labH21.ForeColor = Color.Black;
						this.labH19.ForeColor = Color.Black;
						this.labH19.Text = Hinicia;
						this.labH20.Text = HFinaliza;
						this.labH21.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button7.BackColor = Color.Silver;
					this.labH19.BackColor = Color.Silver;
					this.labH20.BackColor = Color.Silver;
					this.labH21.BackColor = Color.Silver;
					this.labH21.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button7.BackColor = Color.SaddleBrown;
						this.labH19.BackColor = Color.SaddleBrown;
						this.labH20.BackColor = Color.SaddleBrown;
						this.labH21.BackColor = Color.SaddleBrown;
						this.labH19.Text = hri;
						this.labH20.Text = hrf;
						this.labH21.Text = "RESERVADO";
					}
					else
					{
						this.button7.BackColor = Color.Green;
						this.labH19.BackColor = Color.Green;
						this.labH20.BackColor = Color.Green;
						this.labH21.BackColor = Color.Green;
						this.labH19.Text = "00:00";
						this.labH20.Text = "00:00";
						this.labH21.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn7.Visible = true;
					this.labelEn7.BackColor = Color.Red;
					this.labelEn7.ForeColor = Color.White;
					this.labelEn7.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn7.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button7.BackColor = Color.MediumPurple;
					this.labH19.BackColor = Color.MediumPurple;
					this.labH20.BackColor = Color.MediumPurple;
					this.labH21.BackColor = Color.MediumPurple;
					this.labH19.Text = "";
					this.labH20.Text = "";
					this.labH21.Font = new Font(this.labH21.Font.FontFamily, 10);
					this.labH21.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button7.BackColor = Color.DarkSlateBlue;
					this.labH19.BackColor = Color.DarkSlateBlue;
					this.labH20.BackColor = Color.DarkSlateBlue;
					this.labH21.BackColor = Color.DarkSlateBlue;
					this.labH19.Text = "";
					this.labH20.Text = "";
					this.labH21.Font = new Font(this.labH21.Font.FontFamily, 11);
					this.labH21.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button7.BackColor = Color.Turquoise;
					this.labH19.BackColor = Color.Turquoise;
					this.labH20.BackColor = Color.Turquoise;
					this.labH21.BackColor = Color.Turquoise;
					this.button7.ForeColor = Color.Black;
					this.labH19.ForeColor = Color.Black;
					this.labH20.ForeColor = Color.Black;
					this.labH21.ForeColor = Color.Black;
					this.labH19.Text = "";
					this.labH20.Text = "";
					this.labH21.Font = new Font(this.labH21.Font.FontFamily, 10);
					this.labH21.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button7.BackColor = Color.Navy;
					this.labH19.BackColor = Color.Navy;
					this.labH20.BackColor = Color.Navy;
					this.labH21.BackColor = Color.Navy;
					this.button7.ForeColor = Color.White;
					this.labH19.ForeColor = Color.White;
					this.labH20.ForeColor = Color.White;
					this.labH21.ForeColor = Color.White;
					this.labH19.Text = "";
					this.labH20.Text = "";
					this.labH21.Font = new Font(this.labH21.Font.FontFamily, 10);
					this.labH21.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button7.BackColor = Color.OrangeRed;
					this.labH19.BackColor = Color.OrangeRed;
					this.labH20.BackColor = Color.OrangeRed;
					this.labH21.BackColor = Color.OrangeRed;
					this.button7.ForeColor = Color.White;
					this.labH19.ForeColor = Color.White;
					this.labH20.ForeColor = Color.White;
					this.labH21.ForeColor = Color.White;
					this.labH19.Text = "";
					this.labH20.Text = "";
					this.labH21.Font = new Font(this.labH21.Font.FontFamily, 10);
					this.labH21.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button7.BackColor = Color.Orange;
					this.labH19.BackColor = Color.Orange;
					this.labH20.BackColor = Color.Orange;
					this.labH21.BackColor = Color.Orange;
					this.button7.ForeColor = Color.White;
					this.labH19.ForeColor = Color.White;
					this.labH20.ForeColor = Color.White;
					this.labH21.ForeColor = Color.White;
					this.labH19.Text = "";
					this.labH20.Text = "";
					this.labH21.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH21.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h7
			#region h8
			if (NurmdHabitacion==8)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button8.BackColor = Color.SteelBlue;
						this.labH22.BackColor = Color.SteelBlue;
						this.labH23.BackColor = Color.SteelBlue;
						this.labH24.BackColor = Color.SteelBlue;
						this.labH22.Text = Hinicia;
						this.labH23.Text = HFinaliza;
						this.labH24.Text = "PASADO:"+dif_mins.ToString()+" Min.";	
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button8.BackColor = Color.Firebrick;
						this.labH22.BackColor = Color.Firebrick;
						this.labH23.BackColor = Color.Firebrick;
						this.labH24.BackColor = Color.Firebrick;
						this.labH22.Text = Hinicia;
						this.labH23.Text = HFinaliza;
						this.labH24.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button8.BackColor = Color.Orange;
							this.labH22.BackColor = Color.Orange;
							this.labH23.BackColor = Color.Orange;
							this.labH24.BackColor = Color.Orange;
							this.labH22.Text = Hinicia;
							this.labH23.Text = HFinaliza;
							this.labH24.Text = "T. QUEDA";
						}
						else
						{
							this.button8.BackColor = Color.Red;
							this.labH22.BackColor = Color.Red;
							this.labH23.BackColor = Color.Red;
							this.labH24.BackColor = Color.Red;
							this.labH22.Text = Hinicia;
							this.labH23.Text = HFinaliza;
							this.labH24.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{
			
						this.button8.BackColor = Color.Yellow;
						this.button8.ForeColor = Color.Black;
						this.labH22.BackColor = Color.Yellow;
						this.labH23.BackColor = Color.Yellow;
						this.labH24.BackColor = Color.Yellow;
						this.labH22.ForeColor = Color.Black;
						this.labH23.ForeColor = Color.Black;
						this.labH24.ForeColor = Color.Black;
						this.labH22.Text = Hinicia;
						this.labH23.Text = HFinaliza;
						this.labH24.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button8.BackColor = Color.Silver;
					this.labH22.BackColor = Color.Silver;
					this.labH23.BackColor = Color.Silver;
					this.labH24.BackColor = Color.Silver;
					this.labH24.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button8.BackColor = Color.SaddleBrown;
						this.labH22.BackColor = Color.SaddleBrown;
						this.labH23.BackColor = Color.SaddleBrown;
						this.labH24.BackColor = Color.SaddleBrown;
						this.labH22.Text = hri;
						this.labH23.Text = hrf;
						this.labH24.Text = "RESERVADO";
					}
					else
					{
						this.button8.BackColor = Color.Green;
						this.labH22.BackColor = Color.Green;
						this.labH23.BackColor = Color.Green;
						this.labH24.BackColor = Color.Green;
						this.labH22.Text = "00:00";
						this.labH23.Text = "00:00";
						this.labH24.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn8.Visible = true;
					this.labelEn8.BackColor = Color.Red;
					this.labelEn8.ForeColor = Color.White;
					this.labelEn8.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn8.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button8.BackColor = Color.MediumPurple;
					this.labH22.BackColor = Color.MediumPurple;
					this.labH23.BackColor = Color.MediumPurple;
					this.labH24.BackColor = Color.MediumPurple;
					this.labH22.Text = "";
					this.labH23.Text = "";
					this.labH24.Font = new Font(this.labH24.Font.FontFamily, 10);
					this.labH24.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button8.BackColor = Color.DarkSlateBlue;
					this.labH22.BackColor = Color.DarkSlateBlue;
					this.labH23.BackColor = Color.DarkSlateBlue;
					this.labH24.BackColor = Color.DarkSlateBlue;
					this.labH22.Text = "";
					this.labH23.Text = "";
					this.labH24.Font = new Font(this.labH24.Font.FontFamily, 11);
					this.labH24.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button8.BackColor = Color.Turquoise;
					this.labH22.BackColor = Color.Turquoise;
					this.labH23.BackColor = Color.Turquoise;
					this.labH24.BackColor = Color.Turquoise;
					this.button8.ForeColor = Color.Black;
					this.labH22.ForeColor = Color.Black;
					this.labH23.ForeColor = Color.Black;
					this.labH24.ForeColor = Color.Black;
					this.labH22.Text = "";
					this.labH23.Text = "";
					this.labH24.Font = new Font(this.labH24.Font.FontFamily, 10);
					this.labH24.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button8.BackColor = Color.Navy;
					this.labH22.BackColor = Color.Navy;
					this.labH23.BackColor = Color.Navy;
					this.labH24.BackColor = Color.Navy;
					this.button8.ForeColor = Color.White;
					this.labH22.ForeColor = Color.White;
					this.labH23.ForeColor = Color.White;
					this.labH24.ForeColor = Color.White;
					this.labH22.Text = "";
					this.labH23.Text = "";
					this.labH24.Font = new Font(this.labH24.Font.FontFamily, 10);
					this.labH24.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button8.BackColor = Color.OrangeRed;
					this.labH22.BackColor = Color.OrangeRed;
					this.labH23.BackColor = Color.OrangeRed;
					this.labH24.BackColor = Color.OrangeRed;
					this.button8.ForeColor = Color.White;
					this.labH22.ForeColor = Color.White;
					this.labH23.ForeColor = Color.White;
					this.labH24.ForeColor = Color.White;
					this.labH22.Text = "";
					this.labH23.Text = "";
					this.labH24.Font = new Font(this.labH24.Font.FontFamily, 10);
					this.labH24.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button8.BackColor = Color.Orange;
					this.labH22.BackColor = Color.Orange;
					this.labH23.BackColor = Color.Orange;
					this.labH24.BackColor = Color.Orange;
					this.button8.ForeColor = Color.White;
					this.labH22.ForeColor = Color.White;
					this.labH23.ForeColor = Color.White;
					this.labH24.ForeColor = Color.White;
					this.labH22.Text = "";
					this.labH23.Text = "";
					this.labH24.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH24.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h8
			#region h9
			if (NurmdHabitacion==9)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button9.BackColor = Color.SteelBlue;
						this.labH25.BackColor = Color.SteelBlue;
						this.labH26.BackColor = Color.SteelBlue;
						this.labH27.BackColor = Color.SteelBlue;
						this.labH25.Text = Hinicia;
						this.labH26.Text = HFinaliza;
						this.labH27.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button9.BackColor = Color.Firebrick;
						this.labH25.BackColor = Color.Firebrick;
						this.labH26.BackColor = Color.Firebrick;
						this.labH27.BackColor = Color.Firebrick;
						this.labH25.Text = Hinicia;
						this.labH26.Text = HFinaliza;
						this.labH27.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button9.BackColor = Color.Orange;
							this.labH25.BackColor = Color.Orange;
							this.labH26.BackColor = Color.Orange;
							this.labH27.BackColor = Color.Orange;
							this.labH25.Text = Hinicia;
							this.labH26.Text = HFinaliza;
							this.labH27.Text = "T. QUEDA";
						}
						else
						{
							this.button9.BackColor = Color.Red;
							this.labH25.BackColor = Color.Red;
							this.labH26.BackColor = Color.Red;
							this.labH27.BackColor = Color.Red;
							this.labH25.Text = Hinicia;
							this.labH26.Text = HFinaliza;
							this.labH27.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{
					
						this.button9.BackColor = Color.Yellow;
						this.button9.ForeColor = Color.Black;
						this.labH25.BackColor = Color.Yellow;
						this.labH26.BackColor = Color.Yellow;
						this.labH27.BackColor = Color.Yellow;
						this.labH25.ForeColor = Color.Black;
						this.labH26.ForeColor = Color.Black;
						this.labH27.ForeColor = Color.Black;
						this.labH25.Text = Hinicia;
						this.labH26.Text = HFinaliza;
						this.labH27.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button9.BackColor = Color.Silver;
					this.labH25.BackColor = Color.Silver;
					this.labH26.BackColor = Color.Silver;
					this.labH27.BackColor = Color.Silver;
					this.labH27.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button9.BackColor = Color.SaddleBrown;
						this.labH25.BackColor = Color.SaddleBrown;
						this.labH26.BackColor = Color.SaddleBrown;
						this.labH27.BackColor = Color.SaddleBrown;
						this.labH25.Text = hri;
						this.labH26.Text = hrf;
						this.labH27.Text = "RESERVADO";
					}
					else
					{
						this.button9.BackColor = Color.Green;
						this.labH25.BackColor = Color.Green;
						this.labH26.BackColor = Color.Green;
						this.labH27.BackColor = Color.Green;
						this.labH25.Text = "00:00";
						this.labH26.Text = "00:00";
						this.labH27.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn9.Visible = true;
					this.labelEn9.BackColor = Color.Red;
					this.labelEn9.ForeColor = Color.White;
					this.labelEn9.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn9.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button9.BackColor = Color.MediumPurple;
					this.labH25.BackColor = Color.MediumPurple;
					this.labH26.BackColor = Color.MediumPurple;
					this.labH27.BackColor = Color.MediumPurple;
					this.labH25.Text = "";
					this.labH26.Text = "";
					this.labH27.Font = new Font(this.labH27.Font.FontFamily, 10);
					this.labH27.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button9.BackColor = Color.DarkSlateBlue;
					this.labH25.BackColor = Color.DarkSlateBlue;
					this.labH26.BackColor = Color.DarkSlateBlue;
					this.labH27.BackColor = Color.DarkSlateBlue;
					this.labH25.Text = "";
					this.labH26.Text = "";
					this.labH27.Font = new Font(this.labH27.Font.FontFamily, 11);
					this.labH27.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button9.BackColor = Color.Turquoise;
					this.labH25.BackColor = Color.Turquoise;
					this.labH26.BackColor = Color.Turquoise;
					this.labH27.BackColor = Color.Turquoise;
					this.button9.ForeColor = Color.Black;
					this.labH25.ForeColor = Color.Black;
					this.labH26.ForeColor = Color.Black;
					this.labH27.ForeColor = Color.Black;
					this.labH25.Text = "";
					this.labH26.Text = "";
					this.labH27.Font = new Font(this.labH27.Font.FontFamily, 10);
					this.labH27.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button9.BackColor = Color.Navy;
					this.labH25.BackColor = Color.Navy;
					this.labH26.BackColor = Color.Navy;
					this.labH27.BackColor = Color.Navy;
					this.button9.ForeColor = Color.White;
					this.labH25.ForeColor = Color.White;
					this.labH26.ForeColor = Color.White;
					this.labH27.ForeColor = Color.White;
					this.labH25.Text = "";
					this.labH26.Text = "";
					this.labH27.Font = new Font(this.labH27.Font.FontFamily, 10);
					this.labH27.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button9.BackColor = Color.OrangeRed;
					this.labH25.BackColor = Color.OrangeRed;
					this.labH26.BackColor = Color.OrangeRed;
					this.labH27.BackColor = Color.OrangeRed;
					this.button9.ForeColor = Color.White;
					this.labH25.ForeColor = Color.White;
					this.labH26.ForeColor = Color.White;
					this.labH27.ForeColor = Color.White;
					this.labH25.Text = "";
					this.labH26.Text = "";
					this.labH27.Font = new Font(this.labH27.Font.FontFamily, 10);
					this.labH27.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button9.BackColor = Color.Orange;
					this.labH25.BackColor = Color.Orange;
					this.labH26.BackColor = Color.Orange;
					this.labH27.BackColor = Color.Orange;
					this.button9.ForeColor = Color.White;
					this.labH25.ForeColor = Color.White;
					this.labH26.ForeColor = Color.White;
					this.labH27.ForeColor = Color.White;
					this.labH25.Text = "";
					this.labH26.Text = "";
					this.labH27.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH27.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h9
			#region h10
			if (NurmdHabitacion==10)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button10.BackColor = Color.SteelBlue;
						this.labH28.BackColor = Color.SteelBlue;
						this.labH29.BackColor = Color.SteelBlue;
						this.labH30.BackColor = Color.SteelBlue;
						this.labH28.Text = Hinicia;
						this.labH29.Text = HFinaliza;
						this.labH30.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button10.BackColor = Color.Firebrick;
						this.labH28.BackColor = Color.Firebrick;
						this.labH29.BackColor = Color.Firebrick;
						this.labH30.BackColor = Color.Firebrick;
						this.labH28.Text = Hinicia;
						this.labH29.Text = HFinaliza;
						this.labH30.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button10.BackColor = Color.Orange;
							this.labH28.BackColor = Color.Orange;
							this.labH29.BackColor = Color.Orange;
							this.labH30.BackColor = Color.Orange;
							this.labH28.Text = Hinicia;
							this.labH29.Text = HFinaliza;
							this.labH30.Text = "T. QUEDA";
						}
						else
						{
							this.button10.BackColor = Color.Red;
							this.labH28.BackColor = Color.Red;
							this.labH29.BackColor = Color.Red;
							this.labH30.BackColor = Color.Red;
							this.labH28.Text = Hinicia;
							this.labH29.Text = HFinaliza;
							this.labH30.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

						this.button10.BackColor = Color.Yellow;
						this.button10.ForeColor = Color.Black;
						this.labH28.BackColor = Color.Yellow;
						this.labH29.BackColor = Color.Yellow;
						this.labH30.BackColor = Color.Yellow;
						this.labH28.ForeColor = Color.Black;
						this.labH29.ForeColor = Color.Black;
						this.labH30.ForeColor = Color.Black;
						this.labH28.Text = Hinicia;
						this.labH29.Text = HFinaliza;
						this.labH30.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button10.BackColor = Color.Silver;
					this.labH28.BackColor = Color.Silver;
					this.labH29.BackColor = Color.Silver;
					this.labH30.BackColor = Color.Silver;
					this.labH30.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button10.BackColor = Color.SaddleBrown;
						this.labH28.BackColor = Color.SaddleBrown;
						this.labH29.BackColor = Color.SaddleBrown;
						this.labH30.BackColor = Color.SaddleBrown;
						this.labH28.Text = hri;
						this.labH29.Text = hrf;
						this.labH30.Text = "RESERVADO";
					}
					else
					{
						this.button10.BackColor = Color.Green;
						this.labH28.BackColor = Color.Green;
						this.labH29.BackColor = Color.Green;
						this.labH30.BackColor = Color.Green;
						this.labH28.Text = "00:00";
						this.labH29.Text = "00:00";
						this.labH30.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn10.Visible = true;
					this.labelEn10.BackColor = Color.Red;
					this.labelEn10.ForeColor = Color.White;
					this.labelEn10.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn10.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button10.BackColor = Color.MediumPurple;
					this.labH28.BackColor = Color.MediumPurple;
					this.labH29.BackColor = Color.MediumPurple;
					this.labH30.BackColor = Color.MediumPurple;
					this.labH28.Text = "";
					this.labH29.Text = "";
					this.labH30.Font = new Font(this.labH30.Font.FontFamily, 10);
					this.labH30.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button10.BackColor = Color.DarkSlateBlue;
					this.labH28.BackColor = Color.DarkSlateBlue;
					this.labH29.BackColor = Color.DarkSlateBlue;
					this.labH30.BackColor = Color.DarkSlateBlue;
					this.labH28.Text = "";
					this.labH29.Text = "";
					this.labH30.Font = new Font(this.labH30.Font.FontFamily, 11);
					this.labH30.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button10.BackColor = Color.Turquoise;
					this.labH28.BackColor = Color.Turquoise;
					this.labH29.BackColor = Color.Turquoise;
					this.labH30.BackColor = Color.Turquoise;
					this.button10.ForeColor = Color.Black;
					this.labH28.ForeColor = Color.Black;
					this.labH29.ForeColor = Color.Black;
					this.labH30.ForeColor = Color.Black;
					this.labH28.Text = "";
					this.labH29.Text = "";
					this.labH30.Font = new Font(this.labH30.Font.FontFamily, 10);
					this.labH30.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button10.BackColor = Color.Navy;
					this.labH28.BackColor = Color.Navy;
					this.labH29.BackColor = Color.Navy;
					this.labH30.BackColor = Color.Navy;
					this.button10.ForeColor = Color.White;
					this.labH28.ForeColor = Color.White;
					this.labH29.ForeColor = Color.White;
					this.labH30.ForeColor = Color.White;
					this.labH28.Text = "";
					this.labH29.Text = "";
					this.labH30.Font = new Font(this.labH30.Font.FontFamily, 10);
					this.labH30.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button10.BackColor = Color.OrangeRed;
					this.labH28.BackColor = Color.OrangeRed;
					this.labH29.BackColor = Color.OrangeRed;
					this.labH30.BackColor = Color.OrangeRed;
					this.button10.ForeColor = Color.White;
					this.labH28.ForeColor = Color.White;
					this.labH29.ForeColor = Color.White;
					this.labH30.ForeColor = Color.White;
					this.labH28.Text = "";
					this.labH29.Text = "";
					this.labH30.Font = new Font(this.labH30.Font.FontFamily, 10);
					this.labH30.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button10.BackColor = Color.Orange;
					this.labH28.BackColor = Color.Orange;
					this.labH29.BackColor = Color.Orange;
					this.labH30.BackColor = Color.Orange;
					this.button10.ForeColor = Color.White;
					this.labH28.ForeColor = Color.White;
					this.labH29.ForeColor = Color.White;
					this.labH30.ForeColor = Color.White;
					this.labH28.Text = "";
					this.labH29.Text = "";
					this.labH30.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH30.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h10
			#region h11
			if (NurmdHabitacion==11)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button11.BackColor = Color.SteelBlue;
						this.labH31.BackColor = Color.SteelBlue;
						this.labH32.BackColor = Color.SteelBlue;
						this.labH33.BackColor = Color.SteelBlue;
						this.labH31.Text = Hinicia;
						this.labH32.Text = HFinaliza;
						this.labH33.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button11.BackColor = Color.Firebrick;
						this.labH31.BackColor = Color.Firebrick;
						this.labH32.BackColor = Color.Firebrick;
						this.labH33.BackColor = Color.Firebrick;
						this.labH31.Text = Hinicia;
						this.labH32.Text = HFinaliza;
						this.labH33.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button11.BackColor = Color.Orange;
							this.labH31.BackColor = Color.Orange;
							this.labH32.BackColor = Color.Orange;
							this.labH33.BackColor = Color.Orange;
							this.labH31.Text = Hinicia;
							this.labH32.Text = HFinaliza;
							this.labH33.Text = "T. QUEDA";
						}
						else
						{
							this.button11.BackColor = Color.Red;
							this.labH31.BackColor = Color.Red;
							this.labH32.BackColor = Color.Red;
							this.labH33.BackColor = Color.Red;
							this.labH31.Text = Hinicia;
							this.labH32.Text = HFinaliza;
							this.labH33.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{
				
						this.button11.BackColor = Color.Yellow;
						this.button11.ForeColor = Color.Black;
						this.labH31.BackColor = Color.Yellow;
						this.labH32.BackColor = Color.Yellow;
						this.labH33.BackColor = Color.Yellow;
						this.labH31.ForeColor = Color.Black;
						this.labH32.ForeColor = Color.Black;
						this.labH33.ForeColor = Color.Black;
						this.labH31.Text = Hinicia;
						this.labH32.Text = HFinaliza;
						this.labH33.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button11.BackColor = Color.Silver;
					this.labH31.BackColor = Color.Silver;
					this.labH32.BackColor = Color.Silver;
					this.labH33.BackColor = Color.Silver;
					this.labH33.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button11.BackColor = Color.SaddleBrown;
						this.labH31.BackColor = Color.SaddleBrown;
						this.labH32.BackColor = Color.SaddleBrown;
						this.labH33.BackColor = Color.SaddleBrown;
						this.labH31.Text = hri;
						this.labH32.Text = hrf;
						this.labH33.Text = "RESERVADO";
					}
					else
					{
						this.button11.BackColor = Color.Green;
						this.labH31.BackColor = Color.Green;
						this.labH32.BackColor = Color.Green;
						this.labH33.BackColor = Color.Green;
						this.labH31.Text = "00:00";
						this.labH32.Text = "00:00";
						this.labH33.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn11.Visible = true;
					this.labelEn11.BackColor = Color.Red;
					this.labelEn11.ForeColor = Color.White;
					this.labelEn11.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn11.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button11.BackColor = Color.MediumPurple;
					this.labH31.BackColor = Color.MediumPurple;
					this.labH32.BackColor = Color.MediumPurple;
					this.labH33.BackColor = Color.MediumPurple;
					this.labH31.Text = "";
					this.labH32.Text = "";
					this.labH33.Font = new Font(this.labH33.Font.FontFamily, 10);
					this.labH33.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button11.BackColor = Color.DarkSlateBlue;
					this.labH31.BackColor = Color.DarkSlateBlue;
					this.labH32.BackColor = Color.DarkSlateBlue;
					this.labH33.BackColor = Color.DarkSlateBlue;
					this.labH31.Text = "";
					this.labH32.Text = "";
					this.labH33.Font = new Font(this.labH33.Font.FontFamily, 11);
					this.labH33.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button11.BackColor = Color.Turquoise;
					this.labH31.BackColor = Color.Turquoise;
					this.labH32.BackColor = Color.Turquoise;
					this.labH33.BackColor = Color.Turquoise;
					this.button11.ForeColor = Color.Black;
					this.labH31.ForeColor = Color.Black;
					this.labH32.ForeColor = Color.Black;
					this.labH33.ForeColor = Color.Black;
					this.labH31.Text = "";
					this.labH32.Text = "";
					this.labH33.Font = new Font(this.labH33.Font.FontFamily, 10);
					this.labH33.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button11.BackColor = Color.Navy;
					this.labH31.BackColor = Color.Navy;
					this.labH32.BackColor = Color.Navy;
					this.labH33.BackColor = Color.Navy;
					this.button11.ForeColor = Color.White;
					this.labH31.ForeColor = Color.White;
					this.labH32.ForeColor = Color.White;
					this.labH33.ForeColor = Color.White;
					this.labH31.Text = "";
					this.labH32.Text = "";
					this.labH33.Font = new Font(this.labH33.Font.FontFamily, 10);
					this.labH33.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button11.BackColor = Color.OrangeRed;
					this.labH31.BackColor = Color.OrangeRed;
					this.labH32.BackColor = Color.OrangeRed;
					this.labH33.BackColor = Color.OrangeRed;
					this.button11.ForeColor = Color.White;
					this.labH31.ForeColor = Color.White;
					this.labH32.ForeColor = Color.White;
					this.labH33.ForeColor = Color.White;
					this.labH31.Text = "";
					this.labH32.Text = "";
					this.labH33.Font = new Font(this.labH33.Font.FontFamily, 10);
					this.labH33.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button11.BackColor = Color.Orange;
					this.labH31.BackColor = Color.Orange;
					this.labH32.BackColor = Color.Orange;
					this.labH33.BackColor = Color.Orange;
					this.button11.ForeColor = Color.White;
					this.labH31.ForeColor = Color.White;
					this.labH32.ForeColor = Color.White;
					this.labH33.ForeColor = Color.White;
					this.labH31.Text = "";
					this.labH32.Text = "";
					this.labH33.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH33.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h11
			#region h12
			if (NurmdHabitacion==12)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button12.BackColor = Color.SteelBlue;
						this.labH34.BackColor = Color.SteelBlue;
						this.labH35.BackColor = Color.SteelBlue;
						this.labH36.BackColor = Color.SteelBlue;
						this.labH34.Text = Hinicia;
						this.labH35.Text = HFinaliza;
						this.labH36.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button12.BackColor = Color.Firebrick;
						this.labH34.BackColor = Color.Firebrick;
						this.labH35.BackColor = Color.Firebrick;
						this.labH36.BackColor = Color.Firebrick;
						this.labH34.Text = Hinicia;
						this.labH35.Text = HFinaliza;
						this.labH36.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button12.BackColor = Color.Orange;
							this.labH34.BackColor = Color.Orange;
							this.labH35.BackColor = Color.Orange;
							this.labH36.BackColor = Color.Orange;
							this.labH34.Text = Hinicia;
							this.labH35.Text = HFinaliza;
							this.labH36.Text = "T. QUEDA";
						}
						else
						{
							this.button12.BackColor = Color.Red;
							this.labH34.BackColor = Color.Red;
							this.labH35.BackColor = Color.Red;
							this.labH36.BackColor = Color.Red;
							this.labH34.Text = Hinicia;
							this.labH35.Text = HFinaliza;
							this.labH36.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{
					
						this.button12.BackColor = Color.Yellow;
						this.button12.ForeColor = Color.Black;
						this.labH34.BackColor = Color.Yellow;
						this.labH35.BackColor = Color.Yellow;
						this.labH36.BackColor = Color.Yellow;
						this.labH34.ForeColor = Color.Black;
						this.labH35.ForeColor = Color.Black;
						this.labH36.ForeColor = Color.Black;
						this.labH34.Text = Hinicia;
						this.labH35.Text = HFinaliza;
						this.labH36.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button12.BackColor = Color.Silver;
					this.labH34.BackColor = Color.Silver;
					this.labH35.BackColor = Color.Silver;
					this.labH36.BackColor = Color.Silver;
					this.labH36.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button12.BackColor = Color.SaddleBrown;
						this.labH34.BackColor = Color.SaddleBrown;
						this.labH35.BackColor = Color.SaddleBrown;
						this.labH36.BackColor = Color.SaddleBrown;
						this.labH34.Text = hri;
						this.labH35.Text = hrf;
						this.labH36.Text = "RESERVADO";
					}
					else
					{
						this.button12.BackColor = Color.Green;
						this.labH34.BackColor = Color.Green;
						this.labH35.BackColor = Color.Green;
						this.labH36.BackColor = Color.Green;
						this.labH34.Text = "00:00";
						this.labH35.Text = "00:00";
						this.labH36.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn12.Visible = true;
					this.labelEn12.BackColor = Color.Red;
					this.labelEn12.ForeColor = Color.White;
					this.labelEn12.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn12.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button12.BackColor = Color.MediumPurple;
					this.labH34.BackColor = Color.MediumPurple;
					this.labH35.BackColor = Color.MediumPurple;
					this.labH36.BackColor = Color.MediumPurple;
					this.labH34.Text = "";
					this.labH35.Text = "";
					this.labH36.Font = new Font(this.labH36.Font.FontFamily, 10);
					this.labH36.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button12.BackColor = Color.DarkSlateBlue;
					this.labH34.BackColor = Color.DarkSlateBlue;
					this.labH35.BackColor = Color.DarkSlateBlue;
					this.labH36.BackColor = Color.DarkSlateBlue;
					this.labH34.Text = "";
					this.labH35.Text = "";
					this.labH36.Font = new Font(this.labH36.Font.FontFamily, 11);
					this.labH36.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button12.BackColor = Color.Turquoise;
					this.labH34.BackColor = Color.Turquoise;
					this.labH35.BackColor = Color.Turquoise;
					this.labH36.BackColor = Color.Turquoise;
					this.button12.ForeColor = Color.Black;
					this.labH34.ForeColor = Color.Black;
					this.labH35.ForeColor = Color.Black;
					this.labH36.ForeColor = Color.Black;
					this.labH34.Text = "";
					this.labH35.Text = "";
					this.labH36.Font = new Font(this.labH36.Font.FontFamily, 10);
					this.labH36.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button12.BackColor = Color.Navy;
					this.labH34.BackColor = Color.Navy;
					this.labH35.BackColor = Color.Navy;
					this.labH36.BackColor = Color.Navy;
					this.button12.ForeColor = Color.White;
					this.labH34.ForeColor = Color.White;
					this.labH35.ForeColor = Color.White;
					this.labH36.ForeColor = Color.White;
					this.labH34.Text = "";
					this.labH35.Text = "";
					this.labH36.Font = new Font(this.labH36.Font.FontFamily, 10);
					this.labH36.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button12.BackColor = Color.OrangeRed;
					this.labH34.BackColor = Color.OrangeRed;
					this.labH35.BackColor = Color.OrangeRed;
					this.labH36.BackColor = Color.OrangeRed;
					this.button12.ForeColor = Color.White;
					this.labH34.ForeColor = Color.White;
					this.labH35.ForeColor = Color.White;
					this.labH36.ForeColor = Color.White;
					this.labH34.Text = "";
					this.labH35.Text = "";
					this.labH36.Font = new Font(this.labH36.Font.FontFamily, 10);
					this.labH36.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button12.BackColor = Color.Orange;
					this.labH34.BackColor = Color.Orange;
					this.labH35.BackColor = Color.Orange;
					this.labH36.BackColor = Color.Orange;
					this.button12.ForeColor = Color.White;
					this.labH34.ForeColor = Color.White;
					this.labH35.ForeColor = Color.White;
					this.labH36.ForeColor = Color.White;
					this.labH34.Text = "";
					this.labH35.Text = "";
					this.labH36.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH36.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h12
			#region h13
			if (NurmdHabitacion==13)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button13.BackColor = Color.SteelBlue;
						this.labH37.BackColor = Color.SteelBlue;
						this.labH38.BackColor = Color.SteelBlue;
						this.labH39.BackColor = Color.SteelBlue;
						this.labH37.Text = Hinicia;
						this.labH38.Text = HFinaliza;
						this.labH39.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button13.BackColor = Color.Firebrick;
						this.labH37.BackColor = Color.Firebrick;
						this.labH38.BackColor = Color.Firebrick;
						this.labH39.BackColor = Color.Firebrick;
						this.labH37.Text = Hinicia;
						this.labH38.Text = HFinaliza;
						this.labH39.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button13.BackColor = Color.Orange;
							this.labH37.BackColor = Color.Orange;
							this.labH38.BackColor = Color.Orange;
							this.labH39.BackColor = Color.Orange;
							this.labH37.Text = Hinicia;
							this.labH38.Text = HFinaliza;
							this.labH39.Text = "T. QUEDA";
						}
						else
						{
							this.button13.BackColor = Color.Red;
							this.labH37.BackColor = Color.Red;
							this.labH38.BackColor = Color.Red;
							this.labH39.BackColor = Color.Red;
							this.labH37.Text = Hinicia;
							this.labH38.Text = HFinaliza;
							this.labH39.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

						this.button13.BackColor = Color.Yellow;
						this.button13.ForeColor = Color.Black;
						this.labH37.BackColor = Color.Yellow;
						this.labH38.BackColor = Color.Yellow;
						this.labH39.BackColor = Color.Yellow;
						this.labH37.ForeColor = Color.Black;
						this.labH38.ForeColor = Color.Black;
						this.labH39.ForeColor = Color.Black;
						this.labH37.Text = Hinicia;
						this.labH38.Text = HFinaliza;
						this.labH39.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button13.BackColor = Color.Silver;
					this.labH37.BackColor = Color.Silver;
					this.labH38.BackColor = Color.Silver;
					this.labH39.BackColor = Color.Silver;
					this.labH39.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button13.BackColor = Color.SaddleBrown;
						this.labH37.BackColor = Color.SaddleBrown;
						this.labH38.BackColor = Color.SaddleBrown;
						this.labH39.BackColor = Color.SaddleBrown;
						this.labH37.Text = hri;
						this.labH38.Text = hrf;
						this.labH39.Text = "RESERVADO";
					}
					else
					{
						this.button13.BackColor = Color.Green;
						this.labH37.BackColor = Color.Green;
						this.labH38.BackColor = Color.Green;
						this.labH39.BackColor = Color.Green;
						this.labH37.Text = "00:00";
						this.labH38.Text = "00:00";
						this.labH39.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn13.Visible = true;
					this.labelEn13.BackColor = Color.Red;
					this.labelEn13.ForeColor = Color.White;
					this.labelEn13.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn13.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button13.BackColor = Color.MediumPurple;
					this.labH37.BackColor = Color.MediumPurple;
					this.labH38.BackColor = Color.MediumPurple;
					this.labH39.BackColor = Color.MediumPurple;
					this.labH37.Text = "";
					this.labH38.Text = "";
					this.labH39.Font = new Font(this.labH39.Font.FontFamily, 10);
					this.labH39.Text = "INGRESO CLIENTE";

				}
				if (e_hab == 6)
				{
					this.button13.BackColor = Color.DarkSlateBlue;
					this.labH37.BackColor = Color.DarkSlateBlue;
					this.labH38.BackColor = Color.DarkSlateBlue;
					this.labH39.BackColor = Color.DarkSlateBlue;
					this.labH37.Text = "";
					this.labH38.Text = "";
					this.labH39.Font = new Font(this.labH39.Font.FontFamily, 11);
					this.labH39.Text = "ENVÍO COBRO";

				}
				if (e_hab == 7)
				{
					this.button13.BackColor = Color.Turquoise;
					this.labH37.BackColor = Color.Turquoise;
					this.labH38.BackColor = Color.Turquoise;
					this.labH39.BackColor = Color.Turquoise;
					this.button13.ForeColor = Color.Black;
					this.labH37.ForeColor = Color.Black;
					this.labH38.ForeColor = Color.Black;
					this.labH39.ForeColor = Color.Black;
					this.labH37.Text = "";
					this.labH38.Text = "";
					this.labH39.Font = new Font(this.labH39.Font.FontFamily, 10);
					this.labH39.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button13.BackColor = Color.Navy;
					this.labH37.BackColor = Color.Navy;
					this.labH38.BackColor = Color.Navy;
					this.labH39.BackColor = Color.Navy;
					this.button13.ForeColor = Color.White;
					this.labH37.ForeColor = Color.White;
					this.labH38.ForeColor = Color.White;
					this.labH39.ForeColor = Color.White;
					this.labH37.Text = "";
					this.labH38.Text = "";
					this.labH39.Font = new Font(this.labH39.Font.FontFamily, 10);
					this.labH39.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button13.BackColor = Color.OrangeRed;
					this.labH37.BackColor = Color.OrangeRed;
					this.labH38.BackColor = Color.OrangeRed;
					this.labH39.BackColor = Color.OrangeRed;
					this.button13.ForeColor = Color.White;
					this.labH37.ForeColor = Color.White;
					this.labH38.ForeColor = Color.White;
					this.labH39.ForeColor = Color.White;
					this.labH37.Text = "";
					this.labH38.Text = "";
					this.labH39.Font = new Font(this.labH39.Font.FontFamily, 10);
					this.labH39.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button13.BackColor = Color.Orange;
					this.labH37.BackColor = Color.Orange;
					this.labH38.BackColor = Color.Orange;
					this.labH39.BackColor = Color.Orange;
					this.button13.ForeColor = Color.White;
					this.labH37.ForeColor = Color.White;
					this.labH38.ForeColor = Color.White;
					this.labH39.ForeColor = Color.White;
					this.labH37.Text = "";
					this.labH38.Text = "";
					this.labH39.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH39.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h13
			#region h14
			if (NurmdHabitacion==14)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button14.BackColor = Color.SteelBlue;
						this.labH40.BackColor = Color.SteelBlue;
						this.labH41.BackColor = Color.SteelBlue;
						this.labH42.BackColor = Color.SteelBlue;
						this.labH40.Text = Hinicia;
						this.labH41.Text = HFinaliza;
						this.labH42.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button14.BackColor = Color.Firebrick;
						this.labH40.BackColor = Color.Firebrick;
						this.labH41.BackColor = Color.Firebrick;
						this.labH42.BackColor = Color.Firebrick;
						this.labH40.Text = Hinicia;
						this.labH41.Text = HFinaliza;
						this.labH42.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button14.BackColor = Color.Orange;
							this.labH40.BackColor = Color.Orange;
							this.labH41.BackColor = Color.Orange;
							this.labH42.BackColor = Color.Orange;
							this.labH40.Text = Hinicia;
							this.labH41.Text = HFinaliza;
							this.labH42.Text = "T. QUEDA";
						}
						else
						{
							this.button14.BackColor = Color.Red;
							this.labH40.BackColor = Color.Red;
							this.labH41.BackColor = Color.Red;
							this.labH42.BackColor = Color.Red;
							this.labH40.Text = Hinicia;
							this.labH41.Text = HFinaliza;
							this.labH42.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

						this.button14.BackColor = Color.Yellow;
						this.button14.ForeColor = Color.Black;
						this.labH40.BackColor = Color.Yellow;
						this.labH41.BackColor = Color.Yellow;
						this.labH42.BackColor = Color.Yellow;
						this.labH40.ForeColor = Color.Black;
						this.labH41.ForeColor = Color.Black;
						this.labH42.ForeColor = Color.Black;
						this.labH40.Text = Hinicia;
						this.labH41.Text = HFinaliza;
						this.labH42.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button14.BackColor = Color.Silver;
					this.labH40.BackColor = Color.Silver;
					this.labH41.BackColor = Color.Silver;
					this.labH42.BackColor = Color.Silver;
					this.labH42.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button14.BackColor = Color.SaddleBrown;
						this.labH40.BackColor = Color.SaddleBrown;
						this.labH41.BackColor = Color.SaddleBrown;
						this.labH42.BackColor = Color.SaddleBrown;
						this.labH40.Text = hri;
						this.labH41.Text = hrf;
						this.labH42.Text = "RESERVADO";
					}
					else
					{
						this.button14.BackColor = Color.Green;
						this.labH40.BackColor = Color.Green;
						this.labH41.BackColor = Color.Green;
						this.labH42.BackColor = Color.Green;
						this.labH40.Text = "00:00";
						this.labH41.Text = "00:00";
						this.labH42.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn14.Visible = true;
					this.labelEn14.BackColor = Color.Red;
					this.labelEn14.ForeColor = Color.White;
					this.labelEn14.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn14.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button14.BackColor = Color.MediumPurple;
					this.labH40.BackColor = Color.MediumPurple;
					this.labH41.BackColor = Color.MediumPurple;
					this.labH42.BackColor = Color.MediumPurple;
					this.labH40.Text = "";
					this.labH41.Text = "";
					this.labH42.Font = new Font(this.labH42.Font.FontFamily, 10);
					this.labH42.Text = "INGRESO CLIENTE";

				}
				if (e_hab == 6)
				{
					this.button14.BackColor = Color.DarkSlateBlue;
					this.labH40.BackColor = Color.DarkSlateBlue;
					this.labH41.BackColor = Color.DarkSlateBlue;
					this.labH42.BackColor = Color.DarkSlateBlue;
					this.labH40.Text = "";
					this.labH41.Text = "";
					this.labH42.Font = new Font(this.labH42.Font.FontFamily, 11);
					this.labH42.Text = "ENVÍO COBRO";

				}
				if (e_hab == 7)
				{
					this.button14.BackColor = Color.Turquoise;
					this.labH40.BackColor = Color.Turquoise;
					this.labH41.BackColor = Color.Turquoise;
					this.labH42.BackColor = Color.Turquoise;
					this.button14.ForeColor = Color.Black;
					this.labH40.ForeColor = Color.Black;
					this.labH41.ForeColor = Color.Black;
					this.labH42.ForeColor = Color.Black;
					this.labH40.Text = "";
					this.labH41.Text = "";
					this.labH42.Font = new Font(this.labH42.Font.FontFamily, 10);
					this.labH42.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button14.BackColor = Color.Navy;
					this.labH40.BackColor = Color.Navy;
					this.labH41.BackColor = Color.Navy;
					this.labH42.BackColor = Color.Navy;
					this.button14.ForeColor = Color.White;
					this.labH40.ForeColor = Color.White;
					this.labH41.ForeColor = Color.White;
					this.labH42.ForeColor = Color.White;
					this.labH40.Text = "";
					this.labH41.Text = "";
					this.labH42.Font = new Font(this.labH42.Font.FontFamily, 10);
					this.labH42.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button14.BackColor = Color.OrangeRed;
					this.labH40.BackColor = Color.OrangeRed;
					this.labH41.BackColor = Color.OrangeRed;
					this.labH42.BackColor = Color.OrangeRed;
					this.button14.ForeColor = Color.White;
					this.labH40.ForeColor = Color.White;
					this.labH41.ForeColor = Color.White;
					this.labH42.ForeColor = Color.White;
					this.labH40.Text = "";
					this.labH41.Text = "";
					this.labH42.Font = new Font(this.labH42.Font.FontFamily, 10);
					this.labH42.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button14.BackColor = Color.Orange;
					this.labH40.BackColor = Color.Orange;
					this.labH41.BackColor = Color.Orange;
					this.labH42.BackColor = Color.Orange;
					this.button14.ForeColor = Color.White;
					this.labH40.ForeColor = Color.White;
					this.labH41.ForeColor = Color.White;
					this.labH42.ForeColor = Color.White;
					this.labH40.Text = "";
					this.labH41.Text = "";
					this.labH42.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH42.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h14
			#region h15
			if (NurmdHabitacion==15)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button15.BackColor = Color.SteelBlue;
						this.labH43.BackColor = Color.SteelBlue;
						this.labH44.BackColor = Color.SteelBlue;
						this.labH45.BackColor = Color.SteelBlue;
						this.labH43.Text = Hinicia;
						this.labH44.Text = HFinaliza;
						this.labH45.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button15.BackColor = Color.Firebrick;
						this.labH43.BackColor = Color.Firebrick;
						this.labH44.BackColor = Color.Firebrick;
						this.labH45.BackColor = Color.Firebrick;
						this.labH43.Text = Hinicia;
						this.labH44.Text = HFinaliza;
						this.labH45.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button15.BackColor = Color.Orange;
							this.labH43.BackColor = Color.Orange;
							this.labH44.BackColor = Color.Orange;
							this.labH45.BackColor = Color.Orange;
							this.labH43.Text = Hinicia;
							this.labH44.Text = HFinaliza;
							this.labH45.Text = "T. QUEDA";
						}
						else
						{
							this.button15.BackColor = Color.Red;
							this.labH43.BackColor = Color.Red;
							this.labH44.BackColor = Color.Red;
							this.labH45.BackColor = Color.Red;
							this.labH43.Text = Hinicia;
							this.labH44.Text = HFinaliza;
							this.labH45.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

						this.button15.BackColor = Color.Yellow;
						this.button15.ForeColor = Color.Black;
						this.labH43.BackColor = Color.Yellow;
						this.labH44.BackColor = Color.Yellow;
						this.labH45.BackColor = Color.Yellow;
						this.labH43.ForeColor = Color.Black;
						this.labH44.ForeColor = Color.Black;
						this.labH45.ForeColor = Color.Black;
						this.labH43.Text = Hinicia;
						this.labH44.Text = HFinaliza;
						this.labH45.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button15.BackColor = Color.Silver;
					this.labH43.BackColor = Color.Silver;
					this.labH44.BackColor = Color.Silver;
					this.labH45.BackColor = Color.Silver;
					this.labH45.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button15.BackColor = Color.SaddleBrown;
						this.labH43.BackColor = Color.SaddleBrown;
						this.labH44.BackColor = Color.SaddleBrown;
						this.labH45.BackColor = Color.SaddleBrown;
						this.labH43.Text = hri;
						this.labH44.Text = hrf;
						this.labH45.Text = "RESERVADO";
					}
					else
					{
						this.button15.BackColor = Color.Green;
						this.labH43.BackColor = Color.Green;
						this.labH44.BackColor = Color.Green;
						this.labH45.BackColor = Color.Green;
						this.labH43.Text = "00:00";
						this.labH44.Text = "00:00";
						this.labH45.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn15.Visible = true;
					this.labelEn15.BackColor = Color.Red;
					this.labelEn15.ForeColor = Color.White;
					this.labelEn15.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn15.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button15.BackColor = Color.MediumPurple;
					this.labH43.BackColor = Color.MediumPurple;
					this.labH44.BackColor = Color.MediumPurple;
					this.labH45.BackColor = Color.MediumPurple;
					this.labH43.Text = "";
					this.labH44.Text = "";
					this.labH45.Font = new Font(this.labH45.Font.FontFamily, 10);
					this.labH45.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button15.BackColor = Color.DarkSlateBlue;
					this.labH43.BackColor = Color.DarkSlateBlue;
					this.labH44.BackColor = Color.DarkSlateBlue;
					this.labH45.BackColor = Color.DarkSlateBlue;
					this.labH43.Text = "";
					this.labH44.Text = "";
					this.labH45.Font = new Font(this.labH45.Font.FontFamily, 11);
					this.labH45.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button15.BackColor = Color.Turquoise;
					this.labH43.BackColor = Color.Turquoise;
					this.labH44.BackColor = Color.Turquoise;
					this.labH45.BackColor = Color.Turquoise;
					this.button15.ForeColor = Color.Black;
					this.labH43.ForeColor = Color.Black;
					this.labH44.ForeColor = Color.Black;
					this.labH45.ForeColor = Color.Black;
					this.labH43.Text = "";
					this.labH44.Text = "";
					this.labH45.Font = new Font(this.labH45.Font.FontFamily, 10);
					this.labH45.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button15.BackColor = Color.Navy;
					this.labH43.BackColor = Color.Navy;
					this.labH44.BackColor = Color.Navy;
					this.labH45.BackColor = Color.Navy;
					this.button15.ForeColor = Color.White;
					this.labH43.ForeColor = Color.White;
					this.labH44.ForeColor = Color.White;
					this.labH45.ForeColor = Color.White;
					this.labH43.Text = "";
					this.labH44.Text = "";
					this.labH45.Font = new Font(this.labH45.Font.FontFamily, 10);
					this.labH45.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button15.BackColor = Color.OrangeRed;
					this.labH43.BackColor = Color.OrangeRed;
					this.labH44.BackColor = Color.OrangeRed;
					this.labH45.BackColor = Color.OrangeRed;
					this.button15.ForeColor = Color.White;
					this.labH43.ForeColor = Color.White;
					this.labH44.ForeColor = Color.White;
					this.labH45.ForeColor = Color.White;
					this.labH43.Text = "";
					this.labH44.Text = "";
					this.labH45.Font = new Font(this.labH45.Font.FontFamily, 10);
					this.labH45.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button15.BackColor = Color.Orange;
					this.labH43.BackColor = Color.Orange;
					this.labH44.BackColor = Color.Orange;
					this.labH45.BackColor = Color.Orange;
					this.button15.ForeColor = Color.White;
					this.labH43.ForeColor = Color.White;
					this.labH44.ForeColor = Color.White;
					this.labH45.ForeColor = Color.White;
					this.labH43.Text = "";
					this.labH44.Text = "";
					this.labH45.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH45.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h15
			#region h16
			if (NurmdHabitacion==16)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button16.BackColor = Color.SteelBlue;
						this.labH46.BackColor = Color.SteelBlue;
						this.labH47.BackColor = Color.SteelBlue;
						this.labH48.BackColor = Color.SteelBlue;
						this.labH46.Text = Hinicia;
						this.labH47.Text = HFinaliza;
						this.labH48.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button16.BackColor = Color.Firebrick;
						this.labH46.BackColor = Color.Firebrick;
						this.labH47.BackColor = Color.Firebrick;
						this.labH48.BackColor = Color.Firebrick;
						this.labH46.Text = Hinicia;
						this.labH47.Text = HFinaliza;
						this.labH48.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button16.BackColor = Color.Orange;
							this.labH46.BackColor = Color.Orange;
							this.labH47.BackColor = Color.Orange;
							this.labH48.BackColor = Color.Orange;
							this.labH46.Text = Hinicia;
							this.labH47.Text = HFinaliza;
							this.labH48.Text = "T. QUEDA";
						}
						else
						{
							this.button16.BackColor = Color.Red;
							this.labH46.BackColor = Color.Red;
							this.labH47.BackColor = Color.Red;
							this.labH48.BackColor = Color.Red;
							this.labH46.Text = Hinicia;
							this.labH47.Text = HFinaliza;
							this.labH48.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

						this.button16.BackColor = Color.Yellow;
						this.button16.ForeColor = Color.Black;
						this.labH46.BackColor = Color.Yellow;
						this.labH47.BackColor = Color.Yellow;
						this.labH48.BackColor = Color.Yellow;
						this.labH46.BackColor = Color.Black;
						this.labH47.BackColor = Color.Black;
						this.labH48.BackColor = Color.Black;
						this.labH46.Text = Hinicia;
						this.labH47.Text = HFinaliza;
						this.labH48.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button16.BackColor = Color.Silver;
					this.labH46.BackColor = Color.Silver;
					this.labH47.BackColor = Color.Silver;
					this.labH48.BackColor = Color.Silver;
					this.labH48.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button16.BackColor = Color.SaddleBrown;
						this.labH46.BackColor = Color.SaddleBrown;
						this.labH47.BackColor = Color.SaddleBrown;
						this.labH48.BackColor = Color.SaddleBrown;
						this.labH46.Text = hri;
						this.labH47.Text = hrf;
						this.labH48.Text = "RESERVADO";
					}
					else
					{
						this.button16.BackColor = Color.Green;
						this.labH46.BackColor = Color.Green;
						this.labH47.BackColor = Color.Green;
						this.labH48.BackColor = Color.Green;
						this.labH46.Text = "00:00";
						this.labH47.Text = "00:00";
						this.labH48.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn16.Visible = true;
					this.labelEn16.BackColor = Color.Red;
					this.labelEn16.ForeColor = Color.White;
					this.labelEn16.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn16.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button16.BackColor = Color.MediumPurple;
					this.labH46.BackColor = Color.MediumPurple;
					this.labH47.BackColor = Color.MediumPurple;
					this.labH48.BackColor = Color.MediumPurple;
					this.labH46.Text = "";
					this.labH47.Text = "";
					this.labH48.Font = new Font(this.labH48.Font.FontFamily, 10);
					this.labH48.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button16.BackColor = Color.DarkSlateBlue;
					this.labH46.BackColor = Color.DarkSlateBlue;
					this.labH47.BackColor = Color.DarkSlateBlue;
					this.labH48.BackColor = Color.DarkSlateBlue;
					this.labH46.Text = "";
					this.labH47.Text = "";
					this.labH48.Font = new Font(this.labH48.Font.FontFamily, 11);
					this.labH48.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button16.BackColor = Color.Turquoise;
					this.labH46.BackColor = Color.Turquoise;
					this.labH47.BackColor = Color.Turquoise;
					this.labH48.BackColor = Color.Turquoise;
					this.button16.ForeColor = Color.Black;
					this.labH46.ForeColor = Color.Black;
					this.labH47.ForeColor = Color.Black;
					this.labH48.ForeColor = Color.Black;
					this.labH46.Text = "";
					this.labH47.Text = "";
					this.labH48.Font = new Font(this.labH48.Font.FontFamily, 10);
					this.labH48.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button16.BackColor = Color.Navy;
					this.labH46.BackColor = Color.Navy;
					this.labH47.BackColor = Color.Navy;
					this.labH48.BackColor = Color.Navy;
					this.button16.ForeColor = Color.White;
					this.labH46.ForeColor = Color.White;
					this.labH47.ForeColor = Color.White;
					this.labH48.ForeColor = Color.White;
					this.labH46.Text = "";
					this.labH47.Text = "";
					this.labH48.Font = new Font(this.labH48.Font.FontFamily, 10);
					this.labH48.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button16.BackColor = Color.OrangeRed;
					this.labH46.BackColor = Color.OrangeRed;
					this.labH47.BackColor = Color.OrangeRed;
					this.labH48.BackColor = Color.OrangeRed;
					this.button16.ForeColor = Color.White;
					this.labH46.ForeColor = Color.White;
					this.labH47.ForeColor = Color.White;
					this.labH48.ForeColor = Color.White;
					this.labH46.Text = "";
					this.labH47.Text = "";
					this.labH48.Font = new Font(this.labH48.Font.FontFamily, 10);
					this.labH48.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button16.BackColor = Color.Orange;
					this.labH46.BackColor = Color.Orange;
					this.labH47.BackColor = Color.Orange;
					this.labH48.BackColor = Color.Orange;
					this.button16.ForeColor = Color.White;
					this.labH46.ForeColor = Color.White;
					this.labH47.ForeColor = Color.White;
					this.labH48.ForeColor = Color.White;
					this.labH46.Text = "";
					this.labH47.Text = "";
					this.labH48.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH48.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h16
			#region h17
			if (NurmdHabitacion==17)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button17.BackColor = Color.SteelBlue;
						this.labH49.BackColor = Color.SteelBlue;
						this.labH50.BackColor = Color.SteelBlue;
						this.labH51.BackColor = Color.SteelBlue;
						this.labH49.Text = Hinicia;
						this.labH50.Text = HFinaliza;
						this.labH51.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button17.BackColor = Color.Firebrick;
						this.labH49.BackColor = Color.Firebrick;
						this.labH50.BackColor = Color.Firebrick;
						this.labH51.BackColor = Color.Firebrick;
						this.labH49.Text = Hinicia;
						this.labH50.Text = HFinaliza;
						this.labH51.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button17.BackColor = Color.Orange;
							this.labH49.BackColor = Color.Orange;
							this.labH50.BackColor = Color.Orange;
							this.labH51.BackColor = Color.Orange;
							this.labH49.Text = Hinicia;
							this.labH50.Text = HFinaliza;
							this.labH51.Text = "T. QUEDA";
						}
						else
						{
							this.button17.BackColor = Color.Red;
							this.labH49.BackColor = Color.Red;
							this.labH50.BackColor = Color.Red;
							this.labH51.BackColor = Color.Red;
							this.labH49.Text = Hinicia;
							this.labH50.Text = HFinaliza;
							this.labH51.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

						this.button17.BackColor = Color.Yellow;
						this.button17.ForeColor = Color.Black;
						this.labH49.BackColor = Color.Yellow;
						this.labH50.BackColor = Color.Yellow;
						this.labH51.BackColor = Color.Yellow;
						this.labH49.ForeColor = Color.Black;
						this.labH50.ForeColor = Color.Black;
						this.labH51.ForeColor = Color.Black;
						this.labH49.Text = Hinicia;
						this.labH50.Text = HFinaliza;
						this.labH51.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button17.BackColor = Color.Silver;
					this.labH49.BackColor = Color.Silver;
					this.labH50.BackColor = Color.Silver;
					this.labH51.BackColor = Color.Silver;
					this.labH51.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button17.BackColor = Color.SaddleBrown;
						this.labH49.BackColor = Color.SaddleBrown;
						this.labH50.BackColor = Color.SaddleBrown;
						this.labH51.BackColor = Color.SaddleBrown;
						this.labH49.Text = hri;
						this.labH50.Text = hrf;
						this.labH51.Text = "RESERVADO";
					}
					else
					{
						this.button17.BackColor = Color.Green;
						this.labH49.BackColor = Color.Green;
						this.labH50.BackColor = Color.Green;
						this.labH51.BackColor = Color.Green;
						this.labH49.Text = "00:00";
						this.labH50.Text = "00:00";
						this.labH51.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn17.Visible = true;
					this.labelEn17.BackColor = Color.Red;
					this.labelEn17.ForeColor = Color.White;
					this.labelEn17.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn17.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button17.BackColor = Color.MediumPurple;
					this.labH49.BackColor = Color.MediumPurple;
					this.labH50.BackColor = Color.MediumPurple;
					this.labH51.BackColor = Color.MediumPurple;
					this.labH49.Text = "";
					this.labH50.Text = "";
					this.labH51.Font = new Font(this.labH51.Font.FontFamily, 10);
					this.labH51.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button17.BackColor = Color.DarkSlateBlue;
					this.labH49.BackColor = Color.DarkSlateBlue;
					this.labH50.BackColor = Color.DarkSlateBlue;
					this.labH51.BackColor = Color.DarkSlateBlue;
					this.labH49.Text = "";
					this.labH50.Text = "";
					this.labH51.Font = new Font(this.labH51.Font.FontFamily, 11);
					this.labH51.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button17.BackColor = Color.Turquoise;
					this.labH49.BackColor = Color.Turquoise;
					this.labH50.BackColor = Color.Turquoise;
					this.labH51.BackColor = Color.Turquoise;
					this.button17.ForeColor = Color.Black;
					this.labH49.ForeColor = Color.Black;
					this.labH50.ForeColor = Color.Black;
					this.labH51.ForeColor = Color.Black;
					this.labH49.Text = "";
					this.labH50.Text = "";
					this.labH51.Font = new Font(this.labH51.Font.FontFamily, 10);
					this.labH51.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button17.BackColor = Color.Navy;
					this.labH49.BackColor = Color.Navy;
					this.labH50.BackColor = Color.Navy;
					this.labH51.BackColor = Color.Navy;
					this.button17.ForeColor = Color.White;
					this.labH49.ForeColor = Color.White;
					this.labH50.ForeColor = Color.White;
					this.labH51.ForeColor = Color.White;
					this.labH49.Text = "";
					this.labH50.Text = "";
					this.labH51.Font = new Font(this.labH51.Font.FontFamily, 10);
					this.labH51.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button17.BackColor = Color.OrangeRed;
					this.labH49.BackColor = Color.OrangeRed;
					this.labH50.BackColor = Color.OrangeRed;
					this.labH51.BackColor = Color.OrangeRed;
					this.button17.ForeColor = Color.White;
					this.labH49.ForeColor = Color.White;
					this.labH50.ForeColor = Color.White;
					this.labH51.ForeColor = Color.White;
					this.labH49.Text = "";
					this.labH50.Text = "";
					this.labH51.Font = new Font(this.labH51.Font.FontFamily, 10);
					this.labH51.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button17.BackColor = Color.Orange;
					this.labH49.BackColor = Color.Orange;
					this.labH50.BackColor = Color.Orange;
					this.labH51.BackColor = Color.Orange;
					this.button17.ForeColor = Color.White;
					this.labH49.ForeColor = Color.White;
					this.labH50.ForeColor = Color.White;
					this.labH51.ForeColor = Color.White;
					this.labH49.Text = "";
					this.labH50.Text = "";
					this.labH51.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH51.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h17
			#region h18
			if (NurmdHabitacion==18)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button18.BackColor = Color.SteelBlue;
						this.labH52.BackColor = Color.SteelBlue;
						this.labH53.BackColor = Color.SteelBlue;
						this.labH54.BackColor = Color.SteelBlue;
						this.labH52.Text = Hinicia;
						this.labH53.Text = HFinaliza;
						this.labH54.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button18.BackColor = Color.Firebrick;
						this.labH52.BackColor = Color.Firebrick;
						this.labH53.BackColor = Color.Firebrick;
						this.labH54.BackColor = Color.Firebrick;
						this.labH52.Text = Hinicia;
						this.labH53.Text = HFinaliza;
						this.labH54.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button18.BackColor = Color.Orange;
							this.labH52.BackColor = Color.Orange;
							this.labH53.BackColor = Color.Orange;
							this.labH54.BackColor = Color.Orange;
							this.labH52.Text = Hinicia;
							this.labH53.Text = HFinaliza;
							this.labH54.Text = "T. QUEDA";
						}
						else
						{
							this.button18.BackColor = Color.Red;
							this.labH52.BackColor = Color.Red;
							this.labH53.BackColor = Color.Red;
							this.labH54.BackColor = Color.Red;
							this.labH52.Text = Hinicia;
							this.labH53.Text = HFinaliza;
							this.labH54.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{
					
						this.button18.BackColor = Color.Yellow;
						this.button18.ForeColor = Color.Black;
						this.labH52.BackColor = Color.Yellow;
						this.labH53.BackColor = Color.Yellow;
						this.labH54.BackColor = Color.Yellow;
						this.labH52.ForeColor = Color.Black;
						this.labH53.ForeColor = Color.Black;
						this.labH54.ForeColor = Color.Black;
						this.labH52.Text = Hinicia;
						this.labH53.Text = HFinaliza;
						this.labH54.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button18.BackColor = Color.Silver;
					this.labH52.BackColor = Color.Silver;
					this.labH53.BackColor = Color.Silver;
					this.labH54.BackColor = Color.Silver;
					this.labH54.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button18.BackColor = Color.SaddleBrown;
						this.labH52.BackColor = Color.SaddleBrown;
						this.labH53.BackColor = Color.SaddleBrown;
						this.labH54.BackColor = Color.SaddleBrown;
						this.labH52.Text = hri;
						this.labH53.Text = hrf;
						this.labH54.Text = "RESERVADO";
					}
					else
					{
						this.button18.BackColor = Color.Green;
						this.labH52.BackColor = Color.Green;
						this.labH53.BackColor = Color.Green;
						this.labH54.BackColor = Color.Green;
						this.labH52.Text = "00:00";
						this.labH53.Text = "00:00";
						this.labH54.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn18.Visible = true;
					this.labelEn18.BackColor = Color.Red;
					this.labelEn18.ForeColor = Color.White;
					this.labelEn18.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn18.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button18.BackColor = Color.MediumPurple;
					this.labH52.BackColor = Color.MediumPurple;
					this.labH53.BackColor = Color.MediumPurple;
					this.labH54.BackColor = Color.MediumPurple;
					this.labH52.Text = "";
					this.labH53.Text = "";
					this.labH54.Font = new Font(this.labH54.Font.FontFamily, 10);
					this.labH54.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button18.BackColor = Color.DarkSlateBlue;
					this.labH52.BackColor = Color.DarkSlateBlue;
					this.labH53.BackColor = Color.DarkSlateBlue;
					this.labH54.BackColor = Color.DarkSlateBlue;
					this.labH52.Text = "";
					this.labH53.Text = "";
					this.labH54.Font = new Font(this.labH54.Font.FontFamily, 11);
					this.labH54.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button18.BackColor = Color.Turquoise;
					this.labH52.BackColor = Color.Turquoise;
					this.labH53.BackColor = Color.Turquoise;
					this.labH54.BackColor = Color.Turquoise;
					this.button18.ForeColor = Color.Black;
					this.labH52.ForeColor = Color.Black;
					this.labH53.ForeColor = Color.Black;
					this.labH54.ForeColor = Color.Black;
					this.labH52.Text = "";
					this.labH53.Text = "";
					this.labH54.Font = new Font(this.labH54.Font.FontFamily, 10);
					this.labH54.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button18.BackColor = Color.Navy;
					this.labH52.BackColor = Color.Navy;
					this.labH53.BackColor = Color.Navy;
					this.labH54.BackColor = Color.Navy;
					this.button18.ForeColor = Color.White;
					this.labH52.ForeColor = Color.White;
					this.labH53.ForeColor = Color.White;
					this.labH54.ForeColor = Color.White;
					this.labH52.Text = "";
					this.labH53.Text = "";
					this.labH54.Font = new Font(this.labH54.Font.FontFamily, 10);
					this.labH54.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button18.BackColor = Color.OrangeRed;
					this.labH52.BackColor = Color.OrangeRed;
					this.labH53.BackColor = Color.OrangeRed;
					this.labH54.BackColor = Color.OrangeRed;
					this.button18.ForeColor = Color.White;
					this.labH52.ForeColor = Color.White;
					this.labH53.ForeColor = Color.White;
					this.labH54.ForeColor = Color.White;
					this.labH52.Text = "";
					this.labH53.Text = "";
					this.labH54.Font = new Font(this.labH54.Font.FontFamily, 10);
					this.labH54.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button18.BackColor = Color.Orange;
					this.labH52.BackColor = Color.Orange;
					this.labH53.BackColor = Color.Orange;
					this.labH54.BackColor = Color.Orange;
					this.button18.ForeColor = Color.White;
					this.labH52.ForeColor = Color.White;
					this.labH53.ForeColor = Color.White;
					this.labH54.ForeColor = Color.White;
					this.labH52.Text = "";
					this.labH53.Text = "";
					this.labH54.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH54.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h18
			#region h19
			if (NurmdHabitacion==19)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button19.BackColor = Color.SteelBlue;
						this.labH55.BackColor = Color.SteelBlue;
						this.labH56.BackColor = Color.SteelBlue;
						this.labH57.BackColor = Color.SteelBlue;
						this.labH55.Text = Hinicia;
						this.labH56.Text = HFinaliza;
						this.labH57.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button19.BackColor = Color.Firebrick;
						this.labH55.BackColor = Color.Firebrick;
						this.labH56.BackColor = Color.Firebrick;
						this.labH57.BackColor = Color.Firebrick;
						this.labH55.Text = Hinicia;
						this.labH56.Text = HFinaliza;
						this.labH57.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button19.BackColor = Color.Orange;
							this.labH55.BackColor = Color.Orange;
							this.labH56.BackColor = Color.Orange;
							this.labH57.BackColor = Color.Orange;
							this.labH55.Text = Hinicia;
							this.labH56.Text = HFinaliza;
							this.labH57.Text = "T. QUEDA";
						}
						else
						{
							this.button19.BackColor = Color.Red;
							this.labH55.BackColor = Color.Red;
							this.labH56.BackColor = Color.Red;
							this.labH57.BackColor = Color.Red;
							this.labH55.Text = Hinicia;
							this.labH56.Text = HFinaliza;
							this.labH57.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

						this.button19.BackColor = Color.Yellow;
						this.button19.ForeColor = Color.Black;
						this.labH55.BackColor = Color.Yellow;
						this.labH56.BackColor = Color.Yellow;
						this.labH57.BackColor = Color.Yellow;
						this.labH55.ForeColor = Color.Black;
						this.labH56.ForeColor = Color.Black;
						this.labH57.ForeColor = Color.Black;
						this.labH55.Text = Hinicia;
						this.labH56.Text = HFinaliza;
						this.labH57.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button19.BackColor = Color.Silver;
					this.labH55.BackColor = Color.Silver;
					this.labH56.BackColor = Color.Silver;
					this.labH57.BackColor = Color.Silver;
					this.labH57.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button19.BackColor = Color.SaddleBrown;
						this.labH55.BackColor = Color.SaddleBrown;
						this.labH56.BackColor = Color.SaddleBrown;
						this.labH57.BackColor = Color.SaddleBrown;
						this.labH55.Text = hri;
						this.labH56.Text = hrf;
						this.labH57.Text = "RESERVADO";
					}
					else
					{
						this.button19.BackColor = Color.Green;
						this.labH55.BackColor = Color.Green;
						this.labH56.BackColor = Color.Green;
						this.labH57.BackColor = Color.Green;
						this.labH55.Text = "00:00";
						this.labH56.Text = "00:00";
						this.labH57.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn19.Visible = true;
					this.labelEn19.BackColor = Color.Red;
					this.labelEn19.ForeColor = Color.White;
					this.labelEn19.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn19.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button19.BackColor = Color.MediumPurple;
					this.labH55.BackColor = Color.MediumPurple;
					this.labH56.BackColor = Color.MediumPurple;
					this.labH57.BackColor = Color.MediumPurple;
					this.labH55.Text = "";
					this.labH56.Text = "";
					this.labH57.Font = new Font(this.labH57.Font.FontFamily, 10);
					this.labH57.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button19.BackColor = Color.DarkSlateBlue;
					this.labH55.BackColor = Color.DarkSlateBlue;
					this.labH56.BackColor = Color.DarkSlateBlue;
					this.labH57.BackColor = Color.DarkSlateBlue;
					this.labH55.Text = "";
					this.labH56.Text = "";
					this.labH57.Font = new Font(this.labH57.Font.FontFamily, 11);
					this.labH57.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button19.BackColor = Color.Turquoise;
					this.labH55.BackColor = Color.Turquoise;
					this.labH56.BackColor = Color.Turquoise;
					this.labH57.BackColor = Color.Turquoise;
					this.button19.ForeColor = Color.Black;
					this.labH55.ForeColor = Color.Black;
					this.labH56.ForeColor = Color.Black;
					this.labH57.ForeColor = Color.Black;
					this.labH55.Text = "";
					this.labH56.Text = "";
					this.labH57.Font = new Font(this.labH57.Font.FontFamily, 10);
					this.labH57.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button19.BackColor = Color.Navy;
					this.labH55.BackColor = Color.Navy;
					this.labH56.BackColor = Color.Navy;
					this.labH57.BackColor = Color.Navy;
					this.button19.ForeColor = Color.White;
					this.labH55.ForeColor = Color.White;
					this.labH56.ForeColor = Color.White;
					this.labH57.ForeColor = Color.White;
					this.labH55.Text = "";
					this.labH56.Text = "";
					this.labH57.Font = new Font(this.labH57.Font.FontFamily, 10);
					this.labH57.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button19.BackColor = Color.OrangeRed;
					this.labH55.BackColor = Color.OrangeRed;
					this.labH56.BackColor = Color.OrangeRed;
					this.labH57.BackColor = Color.OrangeRed;
					this.button19.ForeColor = Color.White;
					this.labH55.ForeColor = Color.White;
					this.labH56.ForeColor = Color.White;
					this.labH57.ForeColor = Color.White;
					this.labH55.Text = "";
					this.labH56.Text = "";
					this.labH57.Font = new Font(this.labH57.Font.FontFamily, 10);
					this.labH57.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button19.BackColor = Color.Orange;
					this.labH55.BackColor = Color.Orange;
					this.labH56.BackColor = Color.Orange;
					this.labH57.BackColor = Color.Orange;
					this.button19.ForeColor = Color.White;
					this.labH55.ForeColor = Color.White;
					this.labH56.ForeColor = Color.White;
					this.labH57.ForeColor = Color.White;
					this.labH55.Text = "";
					this.labH56.Text = "";
					this.labH57.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH57.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h19
			#region h20
			if (NurmdHabitacion==20)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button20.BackColor = Color.SteelBlue;
						this.labH58.BackColor = Color.SteelBlue;
						this.labH59.BackColor = Color.SteelBlue;
						this.labH60.BackColor = Color.SteelBlue;
						this.labH58.Text = Hinicia;
						this.labH59.Text = HFinaliza;
						this.labH60.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button20.BackColor = Color.Firebrick;
						this.labH58.BackColor = Color.Firebrick;
						this.labH59.BackColor = Color.Firebrick;
						this.labH60.BackColor = Color.Firebrick;
						this.labH58.Text = Hinicia;
						this.labH59.Text = HFinaliza;
						this.labH60.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button20.BackColor = Color.Orange;
							this.labH58.BackColor = Color.Orange;
							this.labH59.BackColor = Color.Orange;
							this.labH60.BackColor = Color.Orange;
							this.labH58.Text = Hinicia;
							this.labH59.Text = HFinaliza;
							this.labH60.Text = "T. QUEDA";
						}
						else
						{
							this.button20.BackColor = Color.Red;
							this.labH58.BackColor = Color.Red;
							this.labH59.BackColor = Color.Red;
							this.labH60.BackColor = Color.Red;
							this.labH58.Text = Hinicia;
							this.labH59.Text = HFinaliza;
							this.labH60.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

						this.button20.BackColor = Color.Yellow;
						this.button20.ForeColor = Color.Black;
						this.labH58.BackColor = Color.Yellow;
						this.labH59.BackColor = Color.Yellow;
						this.labH60.BackColor = Color.Yellow;
						this.labH58.ForeColor = Color.Black;
						this.labH59.ForeColor = Color.Black;
						this.labH60.ForeColor = Color.Black;
						this.labH58.Text = Hinicia;
						this.labH59.Text = HFinaliza;
						this.labH60.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button20.BackColor = Color.Silver;
					this.labH58.BackColor = Color.Silver;
					this.labH59.BackColor = Color.Silver;
					this.labH60.BackColor = Color.Silver;
					this.labH60.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button20.BackColor = Color.SaddleBrown;
						this.labH58.BackColor = Color.SaddleBrown;
						this.labH59.BackColor = Color.SaddleBrown;
						this.labH60.BackColor = Color.SaddleBrown;
						this.labH58.Text = hri;
						this.labH59.Text = hrf;
						this.labH60.Text = "RESERVADO";
					}
					else
					{
						this.button20.BackColor = Color.Green;
						this.labH58.BackColor = Color.Green;
						this.labH59.BackColor = Color.Green;
						this.labH60.BackColor = Color.Green;
						this.labH58.Text = "00:00";
						this.labH59.Text = "00:00";
						this.labH60.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn20.Visible = true;
					this.labelEn20.BackColor = Color.Red;
					this.labelEn20.ForeColor = Color.White;
					this.labelEn20.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn20.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button20.BackColor = Color.MediumPurple;
					this.labH58.BackColor = Color.MediumPurple;
					this.labH59.BackColor = Color.MediumPurple;
					this.labH60.BackColor = Color.MediumPurple;
					this.labH58.Text = "";
					this.labH59.Text = "";
					this.labH60.Font = new Font(this.labH60.Font.FontFamily, 10);
					this.labH60.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button20.BackColor = Color.DarkSlateBlue;
					this.labH58.BackColor = Color.DarkSlateBlue;
					this.labH59.BackColor = Color.DarkSlateBlue;
					this.labH60.BackColor = Color.DarkSlateBlue;
					this.labH58.Text = "";
					this.labH59.Text = "";
					this.labH60.Font = new Font(this.labH60.Font.FontFamily, 11);
					this.labH60.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button20.BackColor = Color.Turquoise;
					this.labH58.BackColor = Color.Turquoise;
					this.labH59.BackColor = Color.Turquoise;
					this.labH60.BackColor = Color.Turquoise;
					this.button20.ForeColor = Color.Black;
					this.labH58.ForeColor = Color.Black;
					this.labH59.ForeColor = Color.Black;
					this.labH60.ForeColor = Color.Black;
					this.labH58.Text = "";
					this.labH59.Text = "";
					this.labH60.Font = new Font(this.labH60.Font.FontFamily, 10);
					this.labH60.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button20.BackColor = Color.Navy;
					this.labH58.BackColor = Color.Navy;
					this.labH59.BackColor = Color.Navy;
					this.labH60.BackColor = Color.Navy;
					this.button20.ForeColor = Color.White;
					this.labH58.ForeColor = Color.White;
					this.labH59.ForeColor = Color.White;
					this.labH60.ForeColor = Color.White;
					this.labH58.Text = "";
					this.labH59.Text = "";
					this.labH60.Font = new Font(this.labH60.Font.FontFamily, 10);
					this.labH60.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button20.BackColor = Color.OrangeRed;
					this.labH58.BackColor = Color.OrangeRed;
					this.labH59.BackColor = Color.OrangeRed;
					this.labH60.BackColor = Color.OrangeRed;
					this.button20.ForeColor = Color.White;
					this.labH58.ForeColor = Color.White;
					this.labH59.ForeColor = Color.White;
					this.labH60.ForeColor = Color.White;
					this.labH58.Text = "";
					this.labH59.Text = "";
					this.labH60.Font = new Font(this.labH60.Font.FontFamily, 10);
					this.labH60.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button20.BackColor = Color.Orange;
					this.labH58.BackColor = Color.Orange;
					this.labH59.BackColor = Color.Orange;
					this.labH60.BackColor = Color.Orange;
					this.button20.ForeColor = Color.White;
					this.labH58.ForeColor = Color.White;
					this.labH59.ForeColor = Color.White;
					this.labH60.ForeColor = Color.White;
					this.labH58.Text = "";
					this.labH59.Text = "";
					this.labH60.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH60.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h20
			#region h21
			if (NurmdHabitacion==21)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button21.BackColor = Color.SteelBlue;
						this.labH61.BackColor = Color.SteelBlue;
						this.labH62.BackColor = Color.SteelBlue;
						this.labH63.BackColor = Color.SteelBlue;
						this.labH61.Text = Hinicia;
						this.labH62.Text = HFinaliza;
						this.labH63.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button21.BackColor = Color.Firebrick;
						this.labH61.BackColor = Color.Firebrick;
						this.labH62.BackColor = Color.Firebrick;
						this.labH63.BackColor = Color.Firebrick;
						this.labH61.Text = Hinicia;
						this.labH62.Text = HFinaliza;
						this.labH63.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button21.BackColor = Color.Orange;
							this.labH61.BackColor = Color.Orange;
							this.labH62.BackColor = Color.Orange;
							this.labH63.BackColor = Color.Orange;
							this.labH61.Text = Hinicia;
							this.labH62.Text = HFinaliza;
							this.labH63.Text = "T. QUEDA";
						}
						else
						{
							this.button21.BackColor = Color.Red;
							this.labH61.BackColor = Color.Red;
							this.labH62.BackColor = Color.Red;
							this.labH63.BackColor = Color.Red;
							this.labH61.Text = Hinicia;
							this.labH62.Text = HFinaliza;
							this.labH63.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

						this.button21.BackColor = Color.Yellow;
						this.button21.ForeColor = Color.Black;
						this.labH61.BackColor = Color.Yellow;
						this.labH62.BackColor = Color.Yellow;
						this.labH63.BackColor = Color.Yellow;
						this.labH61.ForeColor = Color.Black;
						this.labH62.ForeColor = Color.Black;
						this.labH63.ForeColor = Color.Black;
						this.labH61.Text = Hinicia;
						this.labH62.Text = HFinaliza;
						this.labH63.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button21.BackColor = Color.Silver;
					this.labH61.BackColor = Color.Silver;
					this.labH62.BackColor = Color.Silver;
					this.labH63.BackColor = Color.Silver;
					this.labH63.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button21.BackColor = Color.SaddleBrown;
						this.labH61.BackColor = Color.SaddleBrown;
						this.labH62.BackColor = Color.SaddleBrown;
						this.labH63.BackColor = Color.SaddleBrown;
						this.labH61.Text = hri;
						this.labH62.Text = hrf;
						this.labH63.Text = "RESERVADO";
					}
					else
					{
						this.button21.BackColor = Color.Green;
						this.labH61.BackColor = Color.Green;
						this.labH62.BackColor = Color.Green;
						this.labH63.BackColor = Color.Green;
						this.labH61.Text = "00:00";
						this.labH62.Text = "00:00";
						this.labH63.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn21.Visible = true;
					this.labelEn21.BackColor = Color.Red;
					this.labelEn21.ForeColor = Color.White;
					this.labelEn21.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn21.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button21.BackColor = Color.MediumPurple;
					this.labH61.BackColor = Color.MediumPurple;
					this.labH62.BackColor = Color.MediumPurple;
					this.labH63.BackColor = Color.MediumPurple;
					this.labH61.Text = "";
					this.labH62.Text = "";
					this.labH63.Font = new Font(this.labH63.Font.FontFamily, 10);
					this.labH63.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button21.BackColor = Color.DarkSlateBlue;
					this.labH61.BackColor = Color.DarkSlateBlue;
					this.labH62.BackColor = Color.DarkSlateBlue;
					this.labH63.BackColor = Color.DarkSlateBlue;
					this.labH61.Text = "";
					this.labH62.Text = "";
					this.labH63.Font = new Font(this.labH63.Font.FontFamily, 11);
					this.labH63.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button21.BackColor = Color.Turquoise;
					this.labH61.BackColor = Color.Turquoise;
					this.labH62.BackColor = Color.Turquoise;
					this.labH63.BackColor = Color.Turquoise;
					this.button21.ForeColor = Color.Black;
					this.labH61.ForeColor = Color.Black;
					this.labH62.ForeColor = Color.Black;
					this.labH63.ForeColor = Color.Black;
					this.labH61.Text = "";
					this.labH62.Text = "";
					this.labH63.Font = new Font(this.labH63.Font.FontFamily, 10);
					this.labH63.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button21.BackColor = Color.Navy;
					this.labH61.BackColor = Color.Navy;
					this.labH62.BackColor = Color.Navy;
					this.labH63.BackColor = Color.Navy;
					this.button21.ForeColor = Color.White;
					this.labH61.ForeColor = Color.White;
					this.labH62.ForeColor = Color.White;
					this.labH63.ForeColor = Color.White;
					this.labH61.Text = "";
					this.labH62.Text = "";
					this.labH63.Font = new Font(this.labH63.Font.FontFamily, 10);
					this.labH63.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button21.BackColor = Color.OrangeRed;
					this.labH61.BackColor = Color.OrangeRed;
					this.labH62.BackColor = Color.OrangeRed;
					this.labH63.BackColor = Color.OrangeRed;
					this.button21.ForeColor = Color.White;
					this.labH61.ForeColor = Color.White;
					this.labH62.ForeColor = Color.White;
					this.labH63.ForeColor = Color.White;
					this.labH61.Text = "";
					this.labH62.Text = "";
					this.labH63.Font = new Font(this.labH63.Font.FontFamily, 10);
					this.labH63.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button21.BackColor = Color.Orange;
					this.labH61.BackColor = Color.Orange;
					this.labH62.BackColor = Color.Orange;
					this.labH63.BackColor = Color.Orange;
					this.button21.ForeColor = Color.White;
					this.labH61.ForeColor = Color.White;
					this.labH62.ForeColor = Color.White;
					this.labH63.ForeColor = Color.White;
					this.labH61.Text = "";
					this.labH62.Text = "";
					this.labH63.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH63.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h21
			#region h22
			if (NurmdHabitacion==22)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button22.BackColor = Color.SteelBlue;
						this.labH64.BackColor = Color.SteelBlue;
						this.labH65.BackColor = Color.SteelBlue;
						this.labH66.BackColor = Color.SteelBlue;
						this.labH64.Text = Hinicia;
						this.labH65.Text = HFinaliza;
						this.labH66.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button22.BackColor = Color.Firebrick;
						this.labH64.BackColor = Color.Firebrick;
						this.labH65.BackColor = Color.Firebrick;
						this.labH66.BackColor = Color.Firebrick;
						this.labH64.Text = Hinicia;
						this.labH65.Text = HFinaliza;
						this.labH66.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button22.BackColor = Color.Orange;
							this.labH64.BackColor = Color.Orange;
							this.labH65.BackColor = Color.Orange;
							this.labH66.BackColor = Color.Orange;
							this.labH64.Text = Hinicia;
							this.labH65.Text = HFinaliza;
							this.labH66.Text = "T. QUEDA";
						}
						else
						{
							this.button22.BackColor = Color.Red;
							this.labH64.BackColor = Color.Red;
							this.labH65.BackColor = Color.Red;
							this.labH66.BackColor = Color.Red;
							this.labH64.Text = Hinicia;
							this.labH65.Text = HFinaliza;
							this.labH66.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

						this.button22.BackColor = Color.Yellow;
						this.button22.ForeColor = Color.Black;
						this.labH64.BackColor = Color.Yellow;
						this.labH65.BackColor = Color.Yellow;
						this.labH66.BackColor = Color.Yellow;
						this.labH64.ForeColor = Color.Black;
						this.labH65.ForeColor = Color.Black;
						this.labH66.ForeColor = Color.Black;
						this.labH64.Text = Hinicia;
						this.labH65.Text = HFinaliza;
						this.labH66.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button22.BackColor = Color.Silver;
					this.labH64.BackColor = Color.Silver;
					this.labH65.BackColor = Color.Silver;
					this.labH66.BackColor = Color.Silver;
					this.labH66.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button22.BackColor = Color.SaddleBrown;
						this.labH64.BackColor = Color.SaddleBrown;
						this.labH65.BackColor = Color.SaddleBrown;
						this.labH66.BackColor = Color.SaddleBrown;
						this.labH64.Text = hri;
						this.labH65.Text = hrf;
						this.labH66.Text = "RESERVADO";
					}
					else
					{
						this.button22.BackColor = Color.Green;
						this.labH64.BackColor = Color.Green;
						this.labH65.BackColor = Color.Green;
						this.labH66.BackColor = Color.Green;
						this.labH64.Text = "00:00";
						this.labH65.Text = "00:00";
						this.labH66.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn22.Visible = true;
					this.labelEn22.BackColor = Color.Red;
					this.labelEn22.ForeColor = Color.White;
					this.labelEn22.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn22.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button22.BackColor = Color.MediumPurple;
					this.labH64.BackColor = Color.MediumPurple;
					this.labH65.BackColor = Color.MediumPurple;
					this.labH66.BackColor = Color.MediumPurple;
					this.labH64.Text = "";
					this.labH65.Text = "";
					this.labH66.Font = new Font(this.labH66.Font.FontFamily, 10);
					this.labH66.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button22.BackColor = Color.DarkSlateBlue;
					this.labH64.BackColor = Color.DarkSlateBlue;
					this.labH65.BackColor = Color.DarkSlateBlue;
					this.labH66.BackColor = Color.DarkSlateBlue;
					this.labH64.Text = "";
					this.labH65.Text = "";
					this.labH66.Font = new Font(this.labH66.Font.FontFamily, 11);
					this.labH66.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button22.BackColor = Color.Turquoise;
					this.labH64.BackColor = Color.Turquoise;
					this.labH65.BackColor = Color.Turquoise;
					this.labH66.BackColor = Color.Turquoise;
					this.button22.ForeColor = Color.Black;
					this.labH64.ForeColor = Color.Black;
					this.labH65.ForeColor = Color.Black;
					this.labH66.ForeColor = Color.Black;
					this.labH64.Text = "";
					this.labH65.Text = "";
					this.labH66.Font = new Font(this.labH66.Font.FontFamily, 10);
					this.labH66.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button22.BackColor = Color.Navy;
					this.labH64.BackColor = Color.Navy;
					this.labH65.BackColor = Color.Navy;
					this.labH66.BackColor = Color.Navy;
					this.button22.ForeColor = Color.White;
					this.labH64.ForeColor = Color.White;
					this.labH65.ForeColor = Color.White;
					this.labH66.ForeColor = Color.White;
					this.labH64.Text = "";
					this.labH65.Text = "";
					this.labH66.Font = new Font(this.labH66.Font.FontFamily, 10);
					this.labH66.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button22.BackColor = Color.OrangeRed;
					this.labH64.BackColor = Color.OrangeRed;
					this.labH65.BackColor = Color.OrangeRed;
					this.labH66.BackColor = Color.OrangeRed;
					this.button22.ForeColor = Color.White;
					this.labH64.ForeColor = Color.White;
					this.labH65.ForeColor = Color.White;
					this.labH66.ForeColor = Color.White;
					this.labH64.Text = "";
					this.labH65.Text = "";
					this.labH66.Font = new Font(this.labH66.Font.FontFamily, 10);
					this.labH66.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button22.BackColor = Color.Orange;
					this.labH64.BackColor = Color.Orange;
					this.labH65.BackColor = Color.Orange;
					this.labH66.BackColor = Color.Orange;
					this.button22.ForeColor = Color.White;
					this.labH64.ForeColor = Color.White;
					this.labH65.ForeColor = Color.White;
					this.labH66.ForeColor = Color.White;
					this.labH64.Text = "";
					this.labH65.Text = "";
					this.labH66.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH66.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h22
			#region h23
			if (NurmdHabitacion==23)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button23.BackColor = Color.SteelBlue;
						this.labH67.BackColor = Color.SteelBlue;
						this.labH68.BackColor = Color.SteelBlue;
						this.labH69.BackColor = Color.SteelBlue;
						this.labH67.Text = Hinicia;
						this.labH68.Text = HFinaliza;
						this.labH69.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button23.BackColor = Color.Firebrick;
						this.labH67.BackColor = Color.Firebrick;
						this.labH68.BackColor = Color.Firebrick;
						this.labH69.BackColor = Color.Firebrick;
						this.labH67.Text = Hinicia;
						this.labH68.Text = HFinaliza;
						this.labH69.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button23.BackColor = Color.Orange;
							this.labH67.BackColor = Color.Orange;
							this.labH68.BackColor = Color.Orange;
							this.labH69.BackColor = Color.Orange;
							this.labH67.Text = Hinicia;
							this.labH68.Text = HFinaliza;
							this.labH69.Text = "T. QUEDA";
						}
						else
						{
							this.button23.BackColor = Color.Red;
							this.labH67.BackColor = Color.Red;
							this.labH68.BackColor = Color.Red;
							this.labH69.BackColor = Color.Red;
							this.labH67.Text = Hinicia;
							this.labH68.Text = HFinaliza;
							this.labH69.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

						this.button23.BackColor = Color.Yellow;
						this.button23.ForeColor = Color.Black;
						this.labH67.BackColor = Color.Yellow;
						this.labH68.BackColor = Color.Yellow;
						this.labH69.BackColor = Color.Yellow;
						this.labH67.ForeColor = Color.Black;
						this.labH68.ForeColor = Color.Black;
						this.labH69.ForeColor = Color.Black;
						this.labH67.Text = Hinicia;
						this.labH68.Text = HFinaliza;
						this.labH69.Text = "LIMPIEZA";
					
				}
				if (e_hab==4)
				{
					this.button23.BackColor = Color.Silver;
					this.labH67.BackColor = Color.Silver;
					this.labH68.BackColor = Color.Silver;
					this.labH69.BackColor = Color.Silver;
					this.labH69.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button23.BackColor = Color.SaddleBrown;
						this.labH67.BackColor = Color.SaddleBrown;
						this.labH68.BackColor = Color.SaddleBrown;
						this.labH69.BackColor = Color.SaddleBrown;
						this.labH67.Text = hri;
						this.labH68.Text = hrf;
						this.labH69.Text = "RESERVADO";
					}
					else
					{
						this.button23.BackColor = Color.Green;
						this.labH67.BackColor = Color.Green;
						this.labH68.BackColor = Color.Green;
						this.labH69.BackColor = Color.Green;
						this.labH67.Text = "00:00";
						this.labH68.Text = "00:00";
						this.labH69.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn23.Visible = true;
					this.labelEn23.BackColor = Color.Red;
					this.labelEn23.ForeColor = Color.White;
					this.labelEn23.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn23.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button23.BackColor = Color.MediumPurple;
					this.labH67.BackColor = Color.MediumPurple;
					this.labH68.BackColor = Color.MediumPurple;
					this.labH69.BackColor = Color.MediumPurple;
					this.labH67.Text = "";
					this.labH68.Text = "";
					this.labH69.Font = new Font(this.labH69.Font.FontFamily, 10);
					this.labH69.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button23.BackColor = Color.DarkSlateBlue;
					this.labH67.BackColor = Color.DarkSlateBlue;
					this.labH68.BackColor = Color.DarkSlateBlue;
					this.labH69.BackColor = Color.DarkSlateBlue;
					this.labH67.Text = "";
					this.labH68.Text = "";
					this.labH69.Font = new Font(this.labH69.Font.FontFamily, 11);
					this.labH69.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button23.BackColor = Color.Turquoise;
					this.labH67.BackColor = Color.Turquoise;
					this.labH68.BackColor = Color.Turquoise;
					this.labH69.BackColor = Color.Turquoise;
					this.button23.ForeColor = Color.Black;
					this.labH67.ForeColor = Color.Black;
					this.labH68.ForeColor = Color.Black;
					this.labH69.ForeColor = Color.Black;
					this.labH67.Text = "";
					this.labH68.Text = "";
					this.labH69.Font = new Font(this.labH69.Font.FontFamily, 10);
					this.labH69.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button23.BackColor = Color.Navy;
					this.labH67.BackColor = Color.Navy;
					this.labH68.BackColor = Color.Navy;
					this.labH69.BackColor = Color.Navy;
					this.button23.ForeColor = Color.White;
					this.labH67.ForeColor = Color.White;
					this.labH68.ForeColor = Color.White;
					this.labH69.ForeColor = Color.White;
					this.labH67.Text = "";
					this.labH68.Text = "";
					this.labH69.Font = new Font(this.labH69.Font.FontFamily, 10);
					this.labH69.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button23.BackColor = Color.OrangeRed;
					this.labH67.BackColor = Color.OrangeRed;
					this.labH68.BackColor = Color.OrangeRed;
					this.labH69.BackColor = Color.OrangeRed;
					this.button23.ForeColor = Color.White;
					this.labH67.ForeColor = Color.White;
					this.labH68.ForeColor = Color.White;
					this.labH69.ForeColor = Color.White;
					this.labH67.Text = "";
					this.labH68.Text = "";
					this.labH69.Font = new Font(this.labH69.Font.FontFamily, 10);
					this.labH69.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button23.BackColor = Color.Orange;
					this.labH67.BackColor = Color.Orange;
					this.labH68.BackColor = Color.Orange;
					this.labH69.BackColor = Color.Orange;
					this.button23.ForeColor = Color.White;
					this.labH67.ForeColor = Color.White;
					this.labH68.ForeColor = Color.White;
					this.labH69.ForeColor = Color.White;
					this.labH67.Text = "";
					this.labH68.Text = "";
					this.labH69.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH69.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h23
			#region h24
			if (NurmdHabitacion==24)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button24.BackColor = Color.SteelBlue;
						this.labH70.BackColor = Color.SteelBlue;
						this.labH71.BackColor = Color.SteelBlue;
						this.labH72.BackColor = Color.SteelBlue;
						this.labH70.Text = Hinicia;
						this.labH71.Text = HFinaliza;
						this.labH72.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button24.BackColor = Color.Firebrick;
						this.labH70.BackColor = Color.Firebrick;
						this.labH71.BackColor = Color.Firebrick;
						this.labH72.BackColor = Color.Firebrick;
						this.labH70.Text = Hinicia;
						this.labH71.Text = HFinaliza;
						this.labH72.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button24.BackColor = Color.Orange;
							this.labH70.BackColor = Color.Orange;
							this.labH71.BackColor = Color.Orange;
							this.labH72.BackColor = Color.Orange;
							this.labH70.Text = Hinicia;
							this.labH71.Text = HFinaliza;
							this.labH72.Text = "T. QUEDA";
						}
						else
						{
							this.button24.BackColor = Color.Red;
							this.labH70.BackColor = Color.Red;
							this.labH71.BackColor = Color.Red;
							this.labH72.BackColor = Color.Red;
							this.labH70.Text = Hinicia;
							this.labH71.Text = HFinaliza;
							this.labH72.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button24.BackColor = Color.Yellow;
						this.button24.ForeColor = Color.Black;
						this.labH70.BackColor = Color.Yellow;
						this.labH71.BackColor = Color.Yellow;
						this.labH72.BackColor = Color.Yellow;
						this.labH70.ForeColor = Color.Black;
						this.labH71.ForeColor = Color.Black;
						this.labH72.ForeColor = Color.Black;
						this.labH70.Text = Hinicia;
						this.labH71.Text = HFinaliza;
						this.labH72.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button24.BackColor = Color.Silver;
					this.labH70.BackColor = Color.Silver;
					this.labH71.BackColor = Color.Silver;
					this.labH72.BackColor = Color.Silver;
					this.labH72.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button24.BackColor = Color.SaddleBrown;
						this.labH70.BackColor = Color.SaddleBrown;
						this.labH71.BackColor = Color.SaddleBrown;
						this.labH72.BackColor = Color.SaddleBrown;
						this.labH70.Text = hri;
						this.labH71.Text = hrf;
						this.labH72.Text = "RESERVADO";
					}
					else
					{
						this.button24.BackColor = Color.Green;
						this.labH70.BackColor = Color.Green;
						this.labH71.BackColor = Color.Green;
						this.labH72.BackColor = Color.Green;
						this.labH70.Text = "00:00";
						this.labH71.Text = "00:00";
						this.labH72.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn24.Visible = true;
					this.labelEn24.BackColor = Color.Red;
					this.labelEn24.ForeColor = Color.White;
					this.labelEn24.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn24.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button24.BackColor = Color.MediumPurple;
					this.labH70.BackColor = Color.MediumPurple;
					this.labH71.BackColor = Color.MediumPurple;
					this.labH72.BackColor = Color.MediumPurple;
					this.labH70.Text = "";
					this.labH71.Text = "";
					this.labH72.Font = new Font(this.labH72.Font.FontFamily, 10);
					this.labH72.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button24.BackColor = Color.DarkSlateBlue;
					this.labH70.BackColor = Color.DarkSlateBlue;
					this.labH71.BackColor = Color.DarkSlateBlue;
					this.labH72.BackColor = Color.DarkSlateBlue;
					this.labH70.Text = "";
					this.labH71.Text = "";
					this.labH72.Font = new Font(this.labH72.Font.FontFamily, 11);
					this.labH72.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button24.BackColor = Color.Turquoise;
					this.labH70.BackColor = Color.Turquoise;
					this.labH71.BackColor = Color.Turquoise;
					this.labH72.BackColor = Color.Turquoise;
					this.button24.ForeColor = Color.Black;
					this.labH70.ForeColor = Color.Black;
					this.labH71.ForeColor = Color.Black;
					this.labH72.ForeColor = Color.Black;
					this.labH70.Text = "";
					this.labH71.Text = "";
					this.labH72.Font = new Font(this.labH72.Font.FontFamily, 10);
					this.labH72.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button24.BackColor = Color.Navy;
					this.labH70.BackColor = Color.Navy;
					this.labH71.BackColor = Color.Navy;
					this.labH72.BackColor = Color.Navy;
					this.button24.ForeColor = Color.White;
					this.labH70.ForeColor = Color.White;
					this.labH71.ForeColor = Color.White;
					this.labH72.ForeColor = Color.White;
					this.labH70.Text = "";
					this.labH71.Text = "";
					this.labH72.Font = new Font(this.labH72.Font.FontFamily, 10);
					this.labH72.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button24.BackColor = Color.OrangeRed;
					this.labH70.BackColor = Color.OrangeRed;
					this.labH71.BackColor = Color.OrangeRed;
					this.labH72.BackColor = Color.OrangeRed;
					this.button24.ForeColor = Color.White;
					this.labH70.ForeColor = Color.White;
					this.labH71.ForeColor = Color.White;
					this.labH72.ForeColor = Color.White;
					this.labH70.Text = "";
					this.labH71.Text = "";
					this.labH72.Font = new Font(this.labH72.Font.FontFamily, 10);
					this.labH72.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button24.BackColor = Color.Orange;
					this.labH70.BackColor = Color.Orange;
					this.labH71.BackColor = Color.Orange;
					this.labH72.BackColor = Color.Orange;
					this.button24.ForeColor = Color.White;
					this.labH70.ForeColor = Color.White;
					this.labH71.ForeColor = Color.White;
					this.labH72.ForeColor = Color.White;
					this.labH70.Text = "";
					this.labH71.Text = "";
					this.labH72.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH72.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h24
			#region h25
			if (NurmdHabitacion==25)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button25.BackColor = Color.SteelBlue;
						this.labH73.BackColor = Color.SteelBlue;
						this.labH74.BackColor = Color.SteelBlue;
						this.labH75.BackColor = Color.SteelBlue;
						this.labH73.Text = Hinicia;
						this.labH74.Text = HFinaliza;
						this.labH75.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button25.BackColor = Color.Firebrick;
						this.labH73.BackColor = Color.Firebrick;
						this.labH74.BackColor = Color.Firebrick;
						this.labH75.BackColor = Color.Firebrick;
						this.labH73.Text = Hinicia;
						this.labH74.Text = HFinaliza;
						this.labH75.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button25.BackColor = Color.Orange;
							this.labH73.BackColor = Color.Orange;
							this.labH74.BackColor = Color.Orange;
							this.labH75.BackColor = Color.Orange;
							this.labH73.Text = Hinicia;
							this.labH74.Text = HFinaliza;
							this.labH75.Text = "T. QUEDA";
						}
						else
						{
							this.button25.BackColor = Color.Red;
							this.labH73.BackColor = Color.Red;
							this.labH74.BackColor = Color.Red;
							this.labH75.BackColor = Color.Red;
							this.labH73.Text = Hinicia;
							this.labH74.Text = HFinaliza;
							this.labH75.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button25.BackColor = Color.Yellow;
						this.button25.ForeColor = Color.Black;
						this.labH73.BackColor = Color.Yellow;
						this.labH74.BackColor = Color.Yellow;
						this.labH75.BackColor = Color.Yellow;
						this.labH73.ForeColor = Color.Black;
						this.labH74.ForeColor = Color.Black;
						this.labH75.ForeColor = Color.Black;
						this.labH73.Text = Hinicia;
						this.labH74.Text = HFinaliza;
						this.labH75.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button25.BackColor = Color.Silver;
					this.labH73.BackColor = Color.Silver;
					this.labH74.BackColor = Color.Silver;
					this.labH75.BackColor = Color.Silver;
					this.labH75.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button25.BackColor = Color.SaddleBrown;
						this.labH73.BackColor = Color.SaddleBrown;
						this.labH74.BackColor = Color.SaddleBrown;
						this.labH75.BackColor = Color.SaddleBrown;
						this.labH73.Text = hri;
						this.labH74.Text = hrf;
						this.labH75.Text = "RESERVADO";
					}
					else
					{
						this.button25.BackColor = Color.Green;
						this.labH73.BackColor = Color.Green;
						this.labH74.BackColor = Color.Green;
						this.labH75.BackColor = Color.Green;
						this.labH73.Text = "00:00";
						this.labH74.Text = "00:00";
						this.labH75.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn25.Visible = true;
					this.labelEn25.BackColor = Color.Red;
					this.labelEn25.ForeColor = Color.White;
					this.labelEn25.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn25.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button25.BackColor = Color.MediumPurple;
					this.labH73.BackColor = Color.MediumPurple;
					this.labH74.BackColor = Color.MediumPurple;
					this.labH75.BackColor = Color.MediumPurple;
					this.labH73.Text = "";
					this.labH74.Text = "";
					this.labH75.Font = new Font(this.labH75.Font.FontFamily, 10);
					this.labH75.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button25.BackColor = Color.DarkSlateBlue;
					this.labH73.BackColor = Color.DarkSlateBlue;
					this.labH74.BackColor = Color.DarkSlateBlue;
					this.labH75.BackColor = Color.DarkSlateBlue;
					this.labH73.Text = "";
					this.labH74.Text = "";
					this.labH75.Font = new Font(this.labH75.Font.FontFamily, 11);
					this.labH75.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button25.BackColor = Color.Turquoise;
					this.labH73.BackColor = Color.Turquoise;
					this.labH74.BackColor = Color.Turquoise;
					this.labH75.BackColor = Color.Turquoise;
					this.button25.ForeColor = Color.Black;
					this.labH73.ForeColor = Color.Black;
					this.labH74.ForeColor = Color.Black;
					this.labH75.ForeColor = Color.Black;
					this.labH73.Text = "";
					this.labH74.Text = "";
					this.labH75.Font = new Font(this.labH75.Font.FontFamily, 10);
					this.labH75.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button25.BackColor = Color.Navy;
					this.labH73.BackColor = Color.Navy;
					this.labH74.BackColor = Color.Navy;
					this.labH75.BackColor = Color.Navy;
					this.button25.ForeColor = Color.White;
					this.labH73.ForeColor = Color.White;
					this.labH74.ForeColor = Color.White;
					this.labH75.ForeColor = Color.White;
					this.labH73.Text = "";
					this.labH74.Text = "";
					this.labH75.Font = new Font(this.labH75.Font.FontFamily, 10);
					this.labH75.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button25.BackColor = Color.OrangeRed;
					this.labH73.BackColor = Color.OrangeRed;
					this.labH74.BackColor = Color.OrangeRed;
					this.labH75.BackColor = Color.OrangeRed;
					this.button25.ForeColor = Color.White;
					this.labH73.ForeColor = Color.White;
					this.labH74.ForeColor = Color.White;
					this.labH75.ForeColor = Color.White;
					this.labH73.Text = "";
					this.labH74.Text = "";
					this.labH75.Font = new Font(this.labH75.Font.FontFamily, 10);
					this.labH75.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button25.BackColor = Color.Orange;
					this.labH73.BackColor = Color.Orange;
					this.labH74.BackColor = Color.Orange;
					this.labH75.BackColor = Color.Orange;
					this.button25.ForeColor = Color.White;
					this.labH73.ForeColor = Color.White;
					this.labH74.ForeColor = Color.White;
					this.labH75.ForeColor = Color.White;
					this.labH73.Text = "";
					this.labH74.Text = "";
					this.labH75.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH75.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h25
			#region h26
			if (NurmdHabitacion==26)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button26.BackColor = Color.SteelBlue;
						this.labH76.BackColor = Color.SteelBlue;
						this.labH77.BackColor = Color.SteelBlue;
						this.labH78.BackColor = Color.SteelBlue;
						this.labH76.Text = Hinicia;
						this.labH77.Text = HFinaliza;
						this.labH78.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button26.BackColor = Color.Firebrick;
						this.labH76.BackColor = Color.Firebrick;
						this.labH77.BackColor = Color.Firebrick;
						this.labH78.BackColor = Color.Firebrick;
						this.labH76.Text = Hinicia;
						this.labH77.Text = HFinaliza;
						this.labH78.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button26.BackColor = Color.Orange;
							this.labH76.BackColor = Color.Orange;
							this.labH77.BackColor = Color.Orange;
							this.labH78.BackColor = Color.Orange;
							this.labH76.Text = Hinicia;
							this.labH77.Text = HFinaliza;
							this.labH78.Text = "T. QUEDA";
						}
						else
						{
							this.button26.BackColor = Color.Red;
							this.labH76.BackColor = Color.Red;
							this.labH77.BackColor = Color.Red;
							this.labH78.BackColor = Color.Red;
							this.labH76.Text = Hinicia;
							this.labH77.Text = HFinaliza;
							this.labH78.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button26.BackColor = Color.Yellow;
						this.button26.ForeColor = Color.Black;
						this.labH76.BackColor = Color.Yellow;
						this.labH77.BackColor = Color.Yellow;
						this.labH78.BackColor = Color.Yellow;
						this.labH76.ForeColor = Color.Black;
						this.labH77.ForeColor = Color.Black;
						this.labH78.ForeColor = Color.Black;
						this.labH76.Text = Hinicia;
						this.labH77.Text = HFinaliza;
						this.labH78.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button26.BackColor = Color.Silver;
					this.labH76.BackColor = Color.Silver;
					this.labH77.BackColor = Color.Silver;
					this.labH78.BackColor = Color.Silver;
					this.labH78.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button26.BackColor = Color.SaddleBrown;
						this.labH76.BackColor = Color.SaddleBrown;
						this.labH77.BackColor = Color.SaddleBrown;
						this.labH78.BackColor = Color.SaddleBrown;
						this.labH76.Text = hri;
						this.labH77.Text = hrf;
						this.labH78.Text = "RESERVADO";
					}
					else
					{
						this.button26.BackColor = Color.Green;
						this.labH76.BackColor = Color.Green;
						this.labH77.BackColor = Color.Green;
						this.labH78.BackColor = Color.Green;
						this.labH76.Text = "00:00";
						this.labH77.Text = "00:00";
						this.labH78.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn26.Visible = true;
					this.labelEn26.BackColor = Color.Red;
					this.labelEn26.ForeColor = Color.White;
					this.labelEn26.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn26.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button26.BackColor = Color.MediumPurple;
					this.labH76.BackColor = Color.MediumPurple;
					this.labH77.BackColor = Color.MediumPurple;
					this.labH78.BackColor = Color.MediumPurple;
					this.labH76.Text = "";
					this.labH77.Text = "";
					this.labH78.Font = new Font(this.labH78.Font.FontFamily, 10);
					this.labH78.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button26.BackColor = Color.DarkSlateBlue;
					this.labH76.BackColor = Color.DarkSlateBlue;
					this.labH77.BackColor = Color.DarkSlateBlue;
					this.labH78.BackColor = Color.DarkSlateBlue;
					this.labH76.Text = "";
					this.labH77.Text = "";
					this.labH78.Font = new Font(this.labH78.Font.FontFamily, 11);
					this.labH78.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button26.BackColor = Color.Turquoise;
					this.labH76.BackColor = Color.Turquoise;
					this.labH77.BackColor = Color.Turquoise;
					this.labH78.BackColor = Color.Turquoise;
					this.button26.ForeColor = Color.Black;
					this.labH76.ForeColor = Color.Black;
					this.labH77.ForeColor = Color.Black;
					this.labH78.ForeColor = Color.Black;
					this.labH76.Text = "";
					this.labH77.Text = "";
					this.labH78.Font = new Font(this.labH78.Font.FontFamily, 10);
					this.labH78.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button26.BackColor = Color.Navy;
					this.labH76.BackColor = Color.Navy;
					this.labH77.BackColor = Color.Navy;
					this.labH78.BackColor = Color.Navy;
					this.button26.ForeColor = Color.White;
					this.labH76.ForeColor = Color.White;
					this.labH77.ForeColor = Color.White;
					this.labH78.ForeColor = Color.White;
					this.labH76.Text = "";
					this.labH77.Text = "";
					this.labH78.Font = new Font(this.labH78.Font.FontFamily, 10);
					this.labH78.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button26.BackColor = Color.OrangeRed;
					this.labH76.BackColor = Color.OrangeRed;
					this.labH77.BackColor = Color.OrangeRed;
					this.labH78.BackColor = Color.OrangeRed;
					this.button26.ForeColor = Color.White;
					this.labH76.ForeColor = Color.White;
					this.labH77.ForeColor = Color.White;
					this.labH78.ForeColor = Color.White;
					this.labH76.Text = "";
					this.labH77.Text = "";
					this.labH78.Font = new Font(this.labH78.Font.FontFamily, 10);
					this.labH78.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button26.BackColor = Color.Orange;
					this.labH76.BackColor = Color.Orange;
					this.labH77.BackColor = Color.Orange;
					this.labH78.BackColor = Color.Orange;
					this.button26.ForeColor = Color.White;
					this.labH76.ForeColor = Color.White;
					this.labH77.ForeColor = Color.White;
					this.labH78.ForeColor = Color.White;
					this.labH76.Text = "";
					this.labH77.Text = "";
					this.labH78.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH78.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h26
			#region h27
			if (NurmdHabitacion==27)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button27.BackColor = Color.SteelBlue;
						this.labH79.BackColor = Color.SteelBlue;
						this.labH80.BackColor = Color.SteelBlue;
						this.labH81.BackColor = Color.SteelBlue;
						this.labH79.Text = Hinicia;
						this.labH80.Text = HFinaliza;
						this.labH81.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button27.BackColor = Color.Firebrick;
						this.labH79.BackColor = Color.Firebrick;
						this.labH80.BackColor = Color.Firebrick;
						this.labH81.BackColor = Color.Firebrick;
						this.labH79.Text = Hinicia;
						this.labH80.Text = HFinaliza;
						this.labH81.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button27.BackColor = Color.Orange;
							this.labH79.BackColor = Color.Orange;
							this.labH80.BackColor = Color.Orange;
							this.labH81.BackColor = Color.Orange;
							this.labH79.Text = Hinicia;
							this.labH80.Text = HFinaliza;
							this.labH81.Text = "T. QUEDA";
						}
						else
						{
							this.button27.BackColor = Color.Red;
							this.labH79.BackColor = Color.Red;
							this.labH80.BackColor = Color.Red;
							this.labH81.BackColor = Color.Red;
							this.labH79.Text = Hinicia;
							this.labH80.Text = HFinaliza;
							this.labH81.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button27.BackColor = Color.Yellow;
						this.button27.ForeColor = Color.Black;
						this.labH79.BackColor = Color.Yellow;
						this.labH80.BackColor = Color.Yellow;
						this.labH81.BackColor = Color.Yellow;
						this.labH79.ForeColor = Color.Black;
						this.labH80.ForeColor = Color.Black;
						this.labH81.ForeColor = Color.Black;
						this.labH79.Text = Hinicia;
						this.labH80.Text = HFinaliza;
						this.labH81.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button27.BackColor = Color.Silver;
					this.labH79.BackColor = Color.Silver;
					this.labH80.BackColor = Color.Silver;
					this.labH81.BackColor = Color.Silver;
					this.labH81.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button27.BackColor = Color.SaddleBrown;
						this.labH79.BackColor = Color.SaddleBrown;
						this.labH80.BackColor = Color.SaddleBrown;
						this.labH81.BackColor = Color.SaddleBrown;
						this.labH79.Text = hri;
						this.labH80.Text = hrf;
						this.labH81.Text = "RESERVADO";
					}
					else
					{
						this.button27.BackColor = Color.Green;
						this.labH79.BackColor = Color.Green;
						this.labH80.BackColor = Color.Green;
						this.labH81.BackColor = Color.Green;
						this.labH79.Text = "00:00";
						this.labH80.Text = "00:00";
						this.labH81.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn27.Visible = true;
					this.labelEn27.BackColor = Color.Red;
					this.labelEn27.ForeColor = Color.White;
					this.labelEn27.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn27.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button27.BackColor = Color.MediumPurple;
					this.labH79.BackColor = Color.MediumPurple;
					this.labH80.BackColor = Color.MediumPurple;
					this.labH81.BackColor = Color.MediumPurple;
					this.labH79.Text = "";
					this.labH80.Text = "";
					this.labH81.Font = new Font(this.labH81.Font.FontFamily, 10);
					this.labH81.Text = "INGRESO CLIENTE";
				}
				if (e_hab == 6)
				{
					this.button27.BackColor = Color.DarkSlateBlue;
					this.labH79.BackColor = Color.DarkSlateBlue;
					this.labH80.BackColor = Color.DarkSlateBlue;
					this.labH81.BackColor = Color.DarkSlateBlue;
					this.labH79.Text = "";
					this.labH80.Text = "";
					this.labH81.Font = new Font(this.labH81.Font.FontFamily, 11);
					this.labH81.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button27.BackColor = Color.Turquoise;
					this.labH79.BackColor = Color.Turquoise;
					this.labH80.BackColor = Color.Turquoise;
					this.labH81.BackColor = Color.Turquoise;
					this.button27.ForeColor = Color.Black;
					this.labH79.ForeColor = Color.Black;
					this.labH80.ForeColor = Color.Black;
					this.labH81.ForeColor = Color.Black;
					this.labH79.Text = "";
					this.labH80.Text = "";
					this.labH81.Font = new Font(this.labH81.Font.FontFamily, 10);
					this.labH81.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button27.BackColor = Color.Navy;
					this.labH79.BackColor = Color.Navy;
					this.labH80.BackColor = Color.Navy;
					this.labH81.BackColor = Color.Navy;
					this.button27.ForeColor = Color.White;
					this.labH79.ForeColor = Color.White;
					this.labH80.ForeColor = Color.White;
					this.labH81.ForeColor = Color.White;
					this.labH79.Text = "";
					this.labH80.Text = "";
					this.labH81.Font = new Font(this.labH81.Font.FontFamily, 10);
					this.labH81.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button27.BackColor = Color.OrangeRed;
					this.labH79.BackColor = Color.OrangeRed;
					this.labH80.BackColor = Color.OrangeRed;
					this.labH81.BackColor = Color.OrangeRed;
					this.button27.ForeColor = Color.White;
					this.labH79.ForeColor = Color.White;
					this.labH80.ForeColor = Color.White;
					this.labH81.ForeColor = Color.White;
					this.labH79.Text = "";
					this.labH80.Text = "";
					this.labH81.Font = new Font(this.labH81.Font.FontFamily, 10);
					this.labH81.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button27.BackColor = Color.Orange;
					this.labH79.BackColor = Color.Orange;
					this.labH80.BackColor = Color.Orange;
					this.labH81.BackColor = Color.Orange;
					this.button27.ForeColor = Color.White;
					this.labH79.ForeColor = Color.White;
					this.labH80.ForeColor = Color.White;
					this.labH81.ForeColor = Color.White;
					this.labH79.Text = "";
					this.labH80.Text = "";
					this.labH81.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH81.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h27
			#region h28
			if (NurmdHabitacion==28)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button28.BackColor = Color.SteelBlue;
						this.labH82.BackColor = Color.SteelBlue;
						this.labH83.BackColor = Color.SteelBlue;
						this.labH84.BackColor = Color.SteelBlue;
						this.labH82.Text = Hinicia;
						this.labH83.Text = HFinaliza;
						this.labH84.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button28.BackColor = Color.Firebrick;
						this.labH82.BackColor = Color.Firebrick;
						this.labH83.BackColor = Color.Firebrick;
						this.labH84.BackColor = Color.Firebrick;
						this.labH82.Text = Hinicia;
						this.labH83.Text = HFinaliza;
						this.labH84.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button28.BackColor = Color.Orange;
							this.labH82.BackColor = Color.Orange;
							this.labH83.BackColor = Color.Orange;
							this.labH84.BackColor = Color.Orange;
							this.labH82.Text = Hinicia;
							this.labH83.Text = HFinaliza;
							this.labH84.Text = "T. QUEDA";
						}
						else
						{
							this.button28.BackColor = Color.Red;
							this.labH82.BackColor = Color.Red;
							this.labH83.BackColor = Color.Red;
							this.labH84.BackColor = Color.Red;
							this.labH82.Text = Hinicia;
							this.labH83.Text = HFinaliza;
							this.labH84.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button28.BackColor = Color.Yellow;
						this.button28.ForeColor = Color.Black;
						this.labH82.BackColor = Color.Yellow;
						this.labH83.BackColor = Color.Yellow;
						this.labH84.BackColor = Color.Yellow;
						this.labH82.ForeColor = Color.Black;
						this.labH83.ForeColor = Color.Black;
						this.labH84.ForeColor = Color.Black;
						this.labH82.Text = Hinicia;
						this.labH83.Text = HFinaliza;
						this.labH84.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button28.BackColor = Color.Silver;
					this.labH82.BackColor = Color.Silver;
					this.labH83.BackColor = Color.Silver;
					this.labH84.BackColor = Color.Silver;
					this.labH84.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button28.BackColor = Color.SaddleBrown;
						this.labH82.BackColor = Color.SaddleBrown;
						this.labH83.BackColor = Color.SaddleBrown;
						this.labH84.BackColor = Color.SaddleBrown;
						this.labH82.Text = hri;
						this.labH83.Text = hrf;
						this.labH84.Text = "RESERVADO";
					}
					else
					{
						this.button28.BackColor = Color.Green;
						this.labH82.BackColor = Color.Green;
						this.labH83.BackColor = Color.Green;
						this.labH84.BackColor = Color.Green;
						this.labH82.Text = "00:00";
						this.labH83.Text = "00:00";
						this.labH84.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn28.Visible = true;
					this.labelEn28.BackColor = Color.Red;
					this.labelEn28.ForeColor = Color.White;
					this.labelEn28.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn28.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button28.BackColor = Color.MediumPurple;
					this.labH82.BackColor = Color.MediumPurple;
					this.labH83.BackColor = Color.MediumPurple;
					this.labH84.BackColor = Color.MediumPurple;
					this.labH82.Text = "";
					this.labH83.Text = "";
					this.labH84.Font = new Font(this.labH84.Font.FontFamily, 10);
					this.labH84.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button28.BackColor = Color.DarkSlateBlue;
					this.labH82.BackColor = Color.DarkSlateBlue;
					this.labH83.BackColor = Color.DarkSlateBlue;
					this.labH84.BackColor = Color.DarkSlateBlue;
					this.labH82.Text = "";
					this.labH83.Text = "";
					this.labH84.Font = new Font(this.labH84.Font.FontFamily, 11);
					this.labH84.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button28.BackColor = Color.Turquoise;
					this.labH82.BackColor = Color.Turquoise;
					this.labH83.BackColor = Color.Turquoise;
					this.labH84.BackColor = Color.Turquoise;
					this.button28.ForeColor = Color.Black;
					this.labH82.ForeColor = Color.Black;
					this.labH83.ForeColor = Color.Black;
					this.labH84.ForeColor = Color.Black;
					this.labH82.Text = "";
					this.labH83.Text = "";
					this.labH84.Font = new Font(this.labH84.Font.FontFamily, 10);
					this.labH84.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button28.BackColor = Color.Navy;
					this.labH82.BackColor = Color.Navy;
					this.labH83.BackColor = Color.Navy;
					this.labH84.BackColor = Color.Navy;
					this.button28.ForeColor = Color.White;
					this.labH82.ForeColor = Color.White;
					this.labH83.ForeColor = Color.White;
					this.labH84.ForeColor = Color.White;
					this.labH82.Text = "";
					this.labH83.Text = "";
					this.labH84.Font = new Font(this.labH84.Font.FontFamily, 10);
					this.labH84.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button28.BackColor = Color.OrangeRed;
					this.labH82.BackColor = Color.OrangeRed;
					this.labH83.BackColor = Color.OrangeRed;
					this.labH84.BackColor = Color.OrangeRed;
					this.button28.ForeColor = Color.White;
					this.labH82.ForeColor = Color.White;
					this.labH83.ForeColor = Color.White;
					this.labH84.ForeColor = Color.White;
					this.labH82.Text = "";
					this.labH83.Text = "";
					this.labH84.Font = new Font(this.labH84.Font.FontFamily, 10);
					this.labH84.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button28.BackColor = Color.Orange;
					this.labH82.BackColor = Color.Orange;
					this.labH83.BackColor = Color.Orange;
					this.labH84.BackColor = Color.Orange;
					this.button28.ForeColor = Color.White;
					this.labH82.ForeColor = Color.White;
					this.labH83.ForeColor = Color.White;
					this.labH84.ForeColor = Color.White;
					this.labH82.Text = "";
					this.labH83.Text = "";
					this.labH84.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH84.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h28
			#region h29
			if (NurmdHabitacion==29)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button29.BackColor = Color.SteelBlue;
						this.labH85.BackColor = Color.SteelBlue;
						this.labH86.BackColor = Color.SteelBlue;
						this.labH87.BackColor = Color.SteelBlue;
						this.labH85.Text = Hinicia;
						this.labH86.Text = HFinaliza;
						this.labH87.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button29.BackColor = Color.Firebrick;
						this.labH85.BackColor = Color.Firebrick;
						this.labH86.BackColor = Color.Firebrick;
						this.labH87.BackColor = Color.Firebrick;
						this.labH85.Text = Hinicia;
						this.labH86.Text = HFinaliza;
						this.labH87.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button29.BackColor = Color.Orange;
							this.labH85.BackColor = Color.Orange;
							this.labH86.BackColor = Color.Orange;
							this.labH87.BackColor = Color.Orange;
							this.labH85.Text = Hinicia;
							this.labH86.Text = HFinaliza;
							this.labH87.Text = "T. QUEDA";
						}
						else
						{
							this.button29.BackColor = Color.Red;
							this.labH85.BackColor = Color.Red;
							this.labH86.BackColor = Color.Red;
							this.labH87.BackColor = Color.Red;
							this.labH85.Text = Hinicia;
							this.labH86.Text = HFinaliza;
							this.labH87.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button29.BackColor = Color.Yellow;
						this.button29.ForeColor = Color.Black;
						this.labH85.BackColor = Color.Yellow;
						this.labH86.BackColor = Color.Yellow;
						this.labH87.BackColor = Color.Yellow;
						this.labH85.ForeColor = Color.Black;
						this.labH86.ForeColor = Color.Black;
						this.labH87.ForeColor = Color.Black;
						this.labH85.Text = Hinicia;
						this.labH86.Text = HFinaliza;
						this.labH87.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button29.BackColor = Color.Silver;
					this.labH85.BackColor = Color.Silver;
					this.labH86.BackColor = Color.Silver;
					this.labH87.BackColor = Color.Silver;
					this.labH87.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button29.BackColor = Color.SaddleBrown;
						this.labH85.BackColor = Color.SaddleBrown;
						this.labH86.BackColor = Color.SaddleBrown;
						this.labH87.BackColor = Color.SaddleBrown;
						this.labH85.Text = hri;
						this.labH86.Text = hrf;
						this.labH87.Text = "RESERVADO";
					}
					else
					{
						this.button29.BackColor = Color.Green;
						this.labH85.BackColor = Color.Green;
						this.labH86.BackColor = Color.Green;
						this.labH87.BackColor = Color.Green;
						this.labH85.Text = "00:00";
						this.labH86.Text = "00:00";
						this.labH87.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn29.Visible = true;
					this.labelEn29.BackColor = Color.Red;
					this.labelEn29.ForeColor = Color.White;
					this.labelEn29.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn29.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button29.BackColor = Color.MediumPurple;
					this.labH85.BackColor = Color.MediumPurple;
					this.labH86.BackColor = Color.MediumPurple;
					this.labH87.BackColor = Color.MediumPurple;
					this.labH85.Text = "";
					this.labH86.Text = "";
					this.labH87.Font = new Font(this.labH87.Font.FontFamily, 10);
					this.labH87.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button29.BackColor = Color.DarkSlateBlue;
					this.labH85.BackColor = Color.DarkSlateBlue;
					this.labH86.BackColor = Color.DarkSlateBlue;
					this.labH87.BackColor = Color.DarkSlateBlue;
					this.labH85.Text = "";
					this.labH86.Text = "";
					this.labH87.Font = new Font(this.labH87.Font.FontFamily, 11);
					this.labH87.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button29.BackColor = Color.Turquoise;
					this.labH85.BackColor = Color.Turquoise;
					this.labH86.BackColor = Color.Turquoise;
					this.labH87.BackColor = Color.Turquoise;
					this.button29.ForeColor = Color.Black;
					this.labH85.ForeColor = Color.Black;
					this.labH86.ForeColor = Color.Black;
					this.labH87.ForeColor = Color.Black;
					this.labH85.Text = "";
					this.labH86.Text = "";
					this.labH87.Font = new Font(this.labH87.Font.FontFamily, 10);
					this.labH87.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button29.BackColor = Color.Navy;
					this.labH85.BackColor = Color.Navy;
					this.labH86.BackColor = Color.Navy;
					this.labH87.BackColor = Color.Navy;
					this.button29.ForeColor = Color.White;
					this.labH85.ForeColor = Color.White;
					this.labH86.ForeColor = Color.White;
					this.labH87.ForeColor = Color.White;
					this.labH85.Text = "";
					this.labH86.Text = "";
					this.labH87.Font = new Font(this.labH87.Font.FontFamily, 10);
					this.labH87.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button29.BackColor = Color.OrangeRed;
					this.labH85.BackColor = Color.OrangeRed;
					this.labH86.BackColor = Color.OrangeRed;
					this.labH87.BackColor = Color.OrangeRed;
					this.button29.ForeColor = Color.White;
					this.labH85.ForeColor = Color.White;
					this.labH86.ForeColor = Color.White;
					this.labH87.ForeColor = Color.White;
					this.labH85.Text = "";
					this.labH86.Text = "";
					this.labH87.Font = new Font(this.labH87.Font.FontFamily, 10);
					this.labH87.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button29.BackColor = Color.Orange;
					this.labH85.BackColor = Color.Orange;
					this.labH86.BackColor = Color.Orange;
					this.labH87.BackColor = Color.Orange;
					this.button29.ForeColor = Color.White;
					this.labH85.ForeColor = Color.White;
					this.labH86.ForeColor = Color.White;
					this.labH87.ForeColor = Color.White;
					this.labH85.Text = "";
					this.labH86.Text = "";
					this.labH87.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH87.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h29
			#region h30
			if (NurmdHabitacion==30)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button30.BackColor = Color.SteelBlue;
						this.labH88.BackColor = Color.SteelBlue;
						this.labH89.BackColor = Color.SteelBlue;
						this.labH90.BackColor = Color.SteelBlue;
						this.labH88.Text = Hinicia;
						this.labH89.Text = HFinaliza;
						this.labH90.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button30.BackColor = Color.Firebrick;
						this.labH88.BackColor = Color.Firebrick;
						this.labH89.BackColor = Color.Firebrick;
						this.labH90.BackColor = Color.Firebrick;
						this.labH88.Text = Hinicia;
						this.labH89.Text = HFinaliza;
						this.labH90.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button30.BackColor = Color.Orange;
							this.labH88.BackColor = Color.Orange;
							this.labH89.BackColor = Color.Orange;
							this.labH90.BackColor = Color.Orange;
							this.labH88.Text = Hinicia;
							this.labH89.Text = HFinaliza;
							this.labH90.Text = "T. QUEDA";
						}
						else
						{
							this.button30.BackColor = Color.Red;
							this.labH88.BackColor = Color.Red;
							this.labH89.BackColor = Color.Red;
							this.labH90.BackColor = Color.Red;
							this.labH88.Text = Hinicia;
							this.labH89.Text = HFinaliza;
							this.labH90.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button30.BackColor = Color.Yellow;
						this.button30.ForeColor = Color.Black;
						this.labH88.BackColor = Color.Yellow;
						this.labH89.BackColor = Color.Yellow;
						this.labH90.BackColor = Color.Yellow;
						this.labH88.ForeColor = Color.Black;
						this.labH89.ForeColor = Color.Black;
						this.labH90.ForeColor = Color.Black;
						this.labH88.Text = Hinicia;
						this.labH89.Text = HFinaliza;
						this.labH90.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button30.BackColor = Color.Silver;
					this.labH88.BackColor = Color.Silver;
					this.labH89.BackColor = Color.Silver;
					this.labH90.BackColor = Color.Silver;
					this.labH90.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button30.BackColor = Color.SaddleBrown;
						this.labH88.BackColor = Color.SaddleBrown;
						this.labH89.BackColor = Color.SaddleBrown;
						this.labH90.BackColor = Color.SaddleBrown;
						this.labH88.Text = hri;
						this.labH89.Text = hrf;
						this.labH90.Text = "RESERVADO";
					}
					else
					{
						this.button30.BackColor = Color.Green;
						this.labH88.BackColor = Color.Green;
						this.labH89.BackColor = Color.Green;
						this.labH90.BackColor = Color.Green;
						this.labH88.Text = "00:00";
						this.labH89.Text = "00:00";
						this.labH90.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn30.Visible = true;
					this.labelEn30.BackColor = Color.Red;
					this.labelEn30.ForeColor = Color.White;
					this.labelEn30.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn30.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button30.BackColor = Color.MediumPurple;
					this.labH88.BackColor = Color.MediumPurple;
					this.labH89.BackColor = Color.MediumPurple;
					this.labH90.BackColor = Color.MediumPurple;
					this.labH88.Text = "";
					this.labH89.Text = "";
					this.labH90.Font = new Font(this.labH90.Font.FontFamily, 10);
					this.labH90.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button30.BackColor = Color.DarkSlateBlue;
					this.labH88.BackColor = Color.DarkSlateBlue;
					this.labH89.BackColor = Color.DarkSlateBlue;
					this.labH90.BackColor = Color.DarkSlateBlue;
					this.labH88.Text = "";
					this.labH89.Text = "";
					this.labH90.Font = new Font(this.labH90.Font.FontFamily, 11);
					this.labH90.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button30.BackColor = Color.Turquoise;
					this.labH88.BackColor = Color.Turquoise;
					this.labH89.BackColor = Color.Turquoise;
					this.labH90.BackColor = Color.Turquoise;
					this.button30.ForeColor = Color.Black;
					this.labH88.ForeColor = Color.Black;
					this.labH89.ForeColor = Color.Black;
					this.labH90.ForeColor = Color.Black;
					this.labH88.Text = "";
					this.labH89.Text = "";
					this.labH90.Font = new Font(this.labH90.Font.FontFamily, 10);
					this.labH90.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button30.BackColor = Color.Navy;
					this.labH88.BackColor = Color.Navy;
					this.labH89.BackColor = Color.Navy;
					this.labH90.BackColor = Color.Navy;
					this.button30.ForeColor = Color.White;
					this.labH88.ForeColor = Color.White;
					this.labH89.ForeColor = Color.White;
					this.labH90.ForeColor = Color.White;
					this.labH88.Text = "";
					this.labH89.Text = "";
					this.labH90.Font = new Font(this.labH90.Font.FontFamily, 10);
					this.labH90.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button30.BackColor = Color.OrangeRed;
					this.labH88.BackColor = Color.OrangeRed;
					this.labH89.BackColor = Color.OrangeRed;
					this.labH90.BackColor = Color.OrangeRed;
					this.button30.ForeColor = Color.White;
					this.labH88.ForeColor = Color.White;
					this.labH89.ForeColor = Color.White;
					this.labH90.ForeColor = Color.White;
					this.labH88.Text = "";
					this.labH89.Text = "";
					this.labH90.Font = new Font(this.labH90.Font.FontFamily, 10);
					this.labH90.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button30.BackColor = Color.Orange;
					this.labH88.BackColor = Color.Orange;
					this.labH89.BackColor = Color.Orange;
					this.labH90.BackColor = Color.Orange;
					this.button30.ForeColor = Color.White;
					this.labH88.ForeColor = Color.White;
					this.labH89.ForeColor = Color.White;
					this.labH90.ForeColor = Color.White;
					this.labH88.Text = "";
					this.labH89.Text = "";
					this.labH90.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH90.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h30
			#region h31
			if (NurmdHabitacion==31)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button31.BackColor = Color.SteelBlue;
						this.labH91.BackColor = Color.SteelBlue;
						this.labH92.BackColor = Color.SteelBlue;
						this.labH93.BackColor = Color.SteelBlue;
						this.labH91.Text = Hinicia;
						this.labH92.Text = HFinaliza;
						this.labH93.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button31.BackColor = Color.Firebrick;
						this.labH91.BackColor = Color.Firebrick;
						this.labH92.BackColor = Color.Firebrick;
						this.labH93.BackColor = Color.Firebrick;
						this.labH91.Text = Hinicia;
						this.labH92.Text = HFinaliza;
						this.labH93.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button31.BackColor = Color.Orange;
							this.labH91.BackColor = Color.Orange;
							this.labH92.BackColor = Color.Orange;
							this.labH93.BackColor = Color.Orange;
							this.labH91.Text = Hinicia;
							this.labH92.Text = HFinaliza;
							this.labH93.Text = "T. QUEDA";
						}
						else
						{
							this.button31.BackColor = Color.Red;
							this.labH91.BackColor = Color.Red;
							this.labH92.BackColor = Color.Red;
							this.labH93.BackColor = Color.Red;
							this.labH91.Text = Hinicia;
							this.labH92.Text = HFinaliza;
							this.labH93.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button31.BackColor = Color.Yellow;
						this.button31.ForeColor = Color.Black;
						this.labH91.BackColor = Color.Yellow;
						this.labH92.BackColor = Color.Yellow;
						this.labH93.BackColor = Color.Yellow;
						this.labH91.ForeColor = Color.Black;
						this.labH92.ForeColor = Color.Black;
						this.labH93.ForeColor = Color.Black;
						this.labH91.Text = Hinicia;
						this.labH92.Text = HFinaliza;
						this.labH93.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button31.BackColor = Color.Silver;
					this.labH91.BackColor = Color.Silver;
					this.labH92.BackColor = Color.Silver;
					this.labH93.BackColor = Color.Silver;
					this.labH93.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button31.BackColor = Color.SaddleBrown;
						this.labH91.BackColor = Color.SaddleBrown;
						this.labH92.BackColor = Color.SaddleBrown;
						this.labH93.BackColor = Color.SaddleBrown;
						this.labH91.Text = hri;
						this.labH92.Text = hrf;
						this.labH93.Text = "RESERVADO";
					}
					else
					{
						this.button31.BackColor = Color.Green;
						this.labH91.BackColor = Color.Green;
						this.labH92.BackColor = Color.Green;
						this.labH93.BackColor = Color.Green;
						this.labH91.Text = "00:00";
						this.labH92.Text = "00:00";
						this.labH93.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn31.Visible = true;
					this.labelEn31.BackColor = Color.Red;
					this.labelEn31.ForeColor = Color.White;
					this.labelEn31.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn31.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button31.BackColor = Color.MediumPurple;
					this.labH91.BackColor = Color.MediumPurple;
					this.labH92.BackColor = Color.MediumPurple;
					this.labH93.BackColor = Color.MediumPurple;
					this.labH91.Text = "";
					this.labH92.Text = "";
					this.labH93.Font = new Font(this.labH93.Font.FontFamily, 10);
					this.labH93.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button31.BackColor = Color.DarkSlateBlue;
					this.labH91.BackColor = Color.DarkSlateBlue;
					this.labH92.BackColor = Color.DarkSlateBlue;
					this.labH93.BackColor = Color.DarkSlateBlue;
					this.labH91.Text = "";
					this.labH92.Text = "";
					this.labH93.Font = new Font(this.labH93.Font.FontFamily, 11);
					this.labH93.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button31.BackColor = Color.Turquoise;
					this.labH91.BackColor = Color.Turquoise;
					this.labH92.BackColor = Color.Turquoise;
					this.labH93.BackColor = Color.Turquoise;
					this.button31.ForeColor = Color.Black;
					this.labH91.ForeColor = Color.Black;
					this.labH92.ForeColor = Color.Black;
					this.labH93.ForeColor = Color.Black;
					this.labH91.Text = "";
					this.labH92.Text = "";
					this.labH93.Font = new Font(this.labH93.Font.FontFamily, 10);
					this.labH93.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button31.BackColor = Color.Navy;
					this.labH91.BackColor = Color.Navy;
					this.labH92.BackColor = Color.Navy;
					this.labH93.BackColor = Color.Navy;
					this.button31.ForeColor = Color.White;
					this.labH91.ForeColor = Color.White;
					this.labH92.ForeColor = Color.White;
					this.labH93.ForeColor = Color.White;
					this.labH91.Text = "";
					this.labH92.Text = "";
					this.labH93.Font = new Font(this.labH93.Font.FontFamily, 10);
					this.labH93.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button31.BackColor = Color.OrangeRed;
					this.labH91.BackColor = Color.OrangeRed;
					this.labH92.BackColor = Color.OrangeRed;
					this.labH93.BackColor = Color.OrangeRed;
					this.button31.ForeColor = Color.White;
					this.labH91.ForeColor = Color.White;
					this.labH92.ForeColor = Color.White;
					this.labH93.ForeColor = Color.White;
					this.labH91.Text = "";
					this.labH92.Text = "";
					this.labH93.Font = new Font(this.labH93.Font.FontFamily, 10);
					this.labH93.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button31.BackColor = Color.Orange;
					this.labH91.BackColor = Color.Orange;
					this.labH92.BackColor = Color.Orange;
					this.labH93.BackColor = Color.Orange;
					this.button31.ForeColor = Color.White;
					this.labH91.ForeColor = Color.White;
					this.labH92.ForeColor = Color.White;
					this.labH93.ForeColor = Color.White;
					this.labH91.Text = "";
					this.labH92.Text = "";
					this.labH93.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH93.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h31
			#region h32
			if (NurmdHabitacion==32)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button32.BackColor = Color.SteelBlue;
						this.labH94.BackColor = Color.SteelBlue;
						this.labH95.BackColor = Color.SteelBlue;
						this.labH96.BackColor = Color.SteelBlue;
						this.labH94.Text = Hinicia;
						this.labH95.Text = HFinaliza;
						this.labH96.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button32.BackColor = Color.Firebrick;
						this.labH94.BackColor = Color.Firebrick;
						this.labH95.BackColor = Color.Firebrick;
						this.labH96.BackColor = Color.Firebrick;
						this.labH94.Text = Hinicia;
						this.labH95.Text = HFinaliza;
						this.labH96.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button32.BackColor = Color.Orange;
							this.labH94.BackColor = Color.Orange;
							this.labH95.BackColor = Color.Orange;
							this.labH96.BackColor = Color.Orange;
							this.labH94.Text = Hinicia;
							this.labH95.Text = HFinaliza;
							this.labH96.Text = "T. QUEDA";
						}
						else
						{
							this.button32.BackColor = Color.Red;
							this.labH94.BackColor = Color.Red;
							this.labH95.BackColor = Color.Red;
							this.labH96.BackColor = Color.Red;
							this.labH94.Text = Hinicia;
							this.labH95.Text = HFinaliza;
							this.labH96.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button32.BackColor = Color.Yellow;
						this.button32.ForeColor = Color.Black;
						this.labH94.BackColor = Color.Yellow;
						this.labH95.BackColor = Color.Yellow;
						this.labH96.BackColor = Color.Yellow;
						this.labH94.ForeColor = Color.Black;
						this.labH95.ForeColor = Color.Black;
						this.labH96.ForeColor = Color.Black;
						this.labH94.Text = Hinicia;
						this.labH95.Text = HFinaliza;
						this.labH96.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button32.BackColor = Color.Silver;
					this.labH94.BackColor = Color.Silver;
					this.labH95.BackColor = Color.Silver;
					this.labH96.BackColor = Color.Silver;
					this.labH96.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button32.BackColor = Color.SaddleBrown;
						this.labH94.BackColor = Color.SaddleBrown;
						this.labH95.BackColor = Color.SaddleBrown;
						this.labH96.BackColor = Color.SaddleBrown;
						this.labH94.Text = hri;
						this.labH95.Text = hrf;
						this.labH96.Text = "RESERVADO";
					}
					else
					{
						this.button32.BackColor = Color.Green;
						this.labH94.BackColor = Color.Green;
						this.labH95.BackColor = Color.Green;
						this.labH96.BackColor = Color.Green;
						this.labH94.Text = "00:00";
						this.labH95.Text = "00:00";
						this.labH96.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn32.Visible = true;
					this.labelEn32.BackColor = Color.Red;
					this.labelEn32.ForeColor = Color.White;
					this.labelEn32.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn32.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button32.BackColor = Color.MediumPurple;
					this.labH94.BackColor = Color.MediumPurple;
					this.labH95.BackColor = Color.MediumPurple;
					this.labH96.BackColor = Color.MediumPurple;
					this.labH94.Text = "";
					this.labH95.Text = "";
					this.labH96.Font = new Font(this.labH96.Font.FontFamily, 10);
					this.labH96.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button32.BackColor = Color.DarkSlateBlue;
					this.labH94.BackColor = Color.DarkSlateBlue;
					this.labH95.BackColor = Color.DarkSlateBlue;
					this.labH96.BackColor = Color.DarkSlateBlue;
					this.labH94.Text = "";
					this.labH95.Text = "";
					this.labH96.Font = new Font(this.labH96.Font.FontFamily, 11);
					this.labH96.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button32.BackColor = Color.Turquoise;
					this.labH94.BackColor = Color.Turquoise;
					this.labH95.BackColor = Color.Turquoise;
					this.labH96.BackColor = Color.Turquoise;
					this.button32.ForeColor = Color.Black;
					this.labH94.ForeColor = Color.Black;
					this.labH95.ForeColor = Color.Black;
					this.labH96.ForeColor = Color.Black;
					this.labH94.Text = "";
					this.labH95.Text = "";
					this.labH96.Font = new Font(this.labH96.Font.FontFamily, 10);
					this.labH96.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button32.BackColor = Color.Navy;
					this.labH94.BackColor = Color.Navy;
					this.labH95.BackColor = Color.Navy;
					this.labH96.BackColor = Color.Navy;
					this.button32.ForeColor = Color.White;
					this.labH94.ForeColor = Color.White;
					this.labH95.ForeColor = Color.White;
					this.labH96.ForeColor = Color.White;
					this.labH94.Text = "";
					this.labH95.Text = "";
					this.labH96.Font = new Font(this.labH96.Font.FontFamily, 10);
					this.labH96.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button32.BackColor = Color.OrangeRed;
					this.labH94.BackColor = Color.OrangeRed;
					this.labH95.BackColor = Color.OrangeRed;
					this.labH96.BackColor = Color.OrangeRed;
					this.button32.ForeColor = Color.White;
					this.labH94.ForeColor = Color.White;
					this.labH95.ForeColor = Color.White;
					this.labH96.ForeColor = Color.White;
					this.labH94.Text = "";
					this.labH95.Text = "";
					this.labH96.Font = new Font(this.labH96.Font.FontFamily, 10);
					this.labH96.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button32.BackColor = Color.Orange;
					this.labH94.BackColor = Color.Orange;
					this.labH95.BackColor = Color.Orange;
					this.labH96.BackColor = Color.Orange;
					this.button32.ForeColor = Color.White;
					this.labH94.ForeColor = Color.White;
					this.labH95.ForeColor = Color.White;
					this.labH96.ForeColor = Color.White;
					this.labH94.Text = "";
					this.labH95.Text = "";
					this.labH96.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH96.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h32
			#region h33
			if (NurmdHabitacion==33)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button33.BackColor = Color.SteelBlue;
						this.labH97.BackColor = Color.SteelBlue;
						this.labH98.BackColor = Color.SteelBlue;
						this.labH99.BackColor = Color.SteelBlue;
						this.labH97.Text = Hinicia;
						this.labH98.Text = HFinaliza;
						this.labH99.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button33.BackColor = Color.Firebrick;
						this.labH97.BackColor = Color.Firebrick;
						this.labH98.BackColor = Color.Firebrick;
						this.labH99.BackColor = Color.Firebrick;
						this.labH97.Text = Hinicia;
						this.labH98.Text = HFinaliza;
						this.labH99.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button33.BackColor = Color.Orange;
							this.labH97.BackColor = Color.Orange;
							this.labH98.BackColor = Color.Orange;
							this.labH99.BackColor = Color.Orange;
							this.labH97.Text = Hinicia;
							this.labH98.Text = HFinaliza;
							this.labH99.Text = "T. QUEDA";
						}
						else
						{
							this.button33.BackColor = Color.Red;
							this.labH97.BackColor = Color.Red;
							this.labH98.BackColor = Color.Red;
							this.labH99.BackColor = Color.Red;
							this.labH97.Text = Hinicia;
							this.labH98.Text = HFinaliza;
							this.labH99.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button33.BackColor = Color.Yellow;
						this.button33.ForeColor = Color.Black;
						this.labH97.BackColor = Color.Yellow;
						this.labH98.BackColor = Color.Yellow;
						this.labH99.BackColor = Color.Yellow;
						this.labH97.ForeColor = Color.Black;
						this.labH98.ForeColor = Color.Black;
						this.labH99.ForeColor = Color.Black;
						this.labH97.Text = Hinicia;
						this.labH98.Text = HFinaliza;
						this.labH99.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button33.BackColor = Color.Silver;
					this.labH97.BackColor = Color.Silver;
					this.labH98.BackColor = Color.Silver;
					this.labH99.BackColor = Color.Silver;
					this.labH99.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button33.BackColor = Color.SaddleBrown;
						this.labH97.BackColor = Color.SaddleBrown;
						this.labH98.BackColor = Color.SaddleBrown;
						this.labH99.BackColor = Color.SaddleBrown;
						this.labH97.Text = hri;
						this.labH98.Text = hrf;
						this.labH99.Text = "RESERVADO";
					}
					else
					{
						this.button33.BackColor = Color.Green;
						this.labH97.BackColor = Color.Green;
						this.labH98.BackColor = Color.Green;
						this.labH99.BackColor = Color.Green;
						this.labH97.Text = "00:00";
						this.labH98.Text = "00:00";
						this.labH99.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn33.Visible = true;
					this.labelEn33.BackColor = Color.Red;
					this.labelEn33.ForeColor = Color.White;
					this.labelEn33.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn33.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button33.BackColor = Color.MediumPurple;
					this.labH97.BackColor = Color.MediumPurple;
					this.labH98.BackColor = Color.MediumPurple;
					this.labH99.BackColor = Color.MediumPurple;
					this.labH97.Text = "";
					this.labH98.Text = "";
					this.labH99.Font = new Font(this.labH99.Font.FontFamily, 10);
					this.labH99.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button33.BackColor = Color.DarkSlateBlue;
					this.labH97.BackColor = Color.DarkSlateBlue;
					this.labH98.BackColor = Color.DarkSlateBlue;
					this.labH99.BackColor = Color.DarkSlateBlue;
					this.labH97.Text = "";
					this.labH98.Text = "";
					this.labH99.Font = new Font(this.labH99.Font.FontFamily, 11);
					this.labH99.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button33.BackColor = Color.Turquoise;
					this.labH97.BackColor = Color.Turquoise;
					this.labH98.BackColor = Color.Turquoise;
					this.labH99.BackColor = Color.Turquoise;
					this.button33.ForeColor = Color.Black;
					this.labH97.ForeColor = Color.Black;
					this.labH98.ForeColor = Color.Black;
					this.labH99.ForeColor = Color.Black;
					this.labH97.Text = "";
					this.labH98.Text = "";
					this.labH99.Font = new Font(this.labH99.Font.FontFamily, 10);
					this.labH99.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button33.BackColor = Color.Navy;
					this.labH97.BackColor = Color.Navy;
					this.labH98.BackColor = Color.Navy;
					this.labH99.BackColor = Color.Navy;
					this.button33.ForeColor = Color.White;
					this.labH97.ForeColor = Color.White;
					this.labH98.ForeColor = Color.White;
					this.labH99.ForeColor = Color.White;
					this.labH97.Text = "";
					this.labH98.Text = "";
					this.labH99.Font = new Font(this.labH99.Font.FontFamily, 10);
					this.labH99.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button33.BackColor = Color.OrangeRed;
					this.labH97.BackColor = Color.OrangeRed;
					this.labH98.BackColor = Color.OrangeRed;
					this.labH99.BackColor = Color.OrangeRed;
					this.button33.ForeColor = Color.White;
					this.labH97.ForeColor = Color.White;
					this.labH98.ForeColor = Color.White;
					this.labH99.ForeColor = Color.White;
					this.labH97.Text = "";
					this.labH98.Text = "";
					this.labH99.Font = new Font(this.labH99.Font.FontFamily, 10);
					this.labH99.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button33.BackColor = Color.Orange;
					this.labH97.BackColor = Color.Orange;
					this.labH98.BackColor = Color.Orange;
					this.labH99.BackColor = Color.Orange;
					this.button33.ForeColor = Color.White;
					this.labH97.ForeColor = Color.White;
					this.labH98.ForeColor = Color.White;
					this.labH99.ForeColor = Color.White;
					this.labH97.Text = "";
					this.labH98.Text = "";
					this.labH99.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH99.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h33
			#region h34
			if (NurmdHabitacion==34)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button34.BackColor = Color.SteelBlue;
						this.labH100.BackColor = Color.SteelBlue;
						this.labH101.BackColor = Color.SteelBlue;
						this.labH102.BackColor = Color.SteelBlue;
						this.labH100.Text = Hinicia;
						this.labH101.Text = HFinaliza;
						this.labH102.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button34.BackColor = Color.Firebrick;
						this.labH100.BackColor = Color.Firebrick;
						this.labH101.BackColor = Color.Firebrick;
						this.labH102.BackColor = Color.Firebrick;
						this.labH100.Text = Hinicia;
						this.labH101.Text = HFinaliza;
						this.labH102.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button34.BackColor = Color.Orange;
							this.labH100.BackColor = Color.Orange;
							this.labH101.BackColor = Color.Orange;
							this.labH102.BackColor = Color.Orange;
							this.labH100.Text = Hinicia;
							this.labH101.Text = HFinaliza;
							this.labH102.Text = "T. QUEDA";
						}
						else
						{
							this.button34.BackColor = Color.Red;
							this.labH100.BackColor = Color.Red;
							this.labH101.BackColor = Color.Red;
							this.labH102.BackColor = Color.Red;
							this.labH100.Text = Hinicia;
							this.labH101.Text = HFinaliza;
							this.labH102.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button34.BackColor = Color.Yellow;
						this.button34.ForeColor = Color.Black;
						this.labH100.BackColor = Color.Yellow;
						this.labH101.BackColor = Color.Yellow;
						this.labH102.BackColor = Color.Yellow;
						this.labH100.ForeColor = Color.Black;
						this.labH101.ForeColor = Color.Black;
						this.labH102.ForeColor = Color.Black;
						this.labH100.Text = Hinicia;
						this.labH101.Text = HFinaliza;
						this.labH102.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button34.BackColor = Color.Silver;
					this.labH100.BackColor = Color.Silver;
					this.labH101.BackColor = Color.Silver;
					this.labH102.BackColor = Color.Silver;
					this.labH102.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button34.BackColor = Color.SaddleBrown;
						this.labH100.BackColor = Color.SaddleBrown;
						this.labH101.BackColor = Color.SaddleBrown;
						this.labH102.BackColor = Color.SaddleBrown;
						this.labH100.Text = hri;
						this.labH101.Text = hrf;
						this.labH102.Text = "RESERVADO";
					}
					else
					{
						this.button34.BackColor = Color.Green;
						this.labH100.BackColor = Color.Green;
						this.labH101.BackColor = Color.Green;
						this.labH102.BackColor = Color.Green;
						this.labH100.Text = "00:00";
						this.labH101.Text = "00:00";
						this.labH102.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn34.Visible = true;
					this.labelEn34.BackColor = Color.Red;
					this.labelEn34.ForeColor = Color.White;
					this.labelEn34.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn34.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button34.BackColor = Color.MediumPurple;
					this.labH100.BackColor = Color.MediumPurple;
					this.labH101.BackColor = Color.MediumPurple;
					this.labH102.BackColor = Color.MediumPurple;
					this.labH100.Text = "";
					this.labH101.Text = "";
					this.labH102.Font = new Font(this.labH102.Font.FontFamily, 10);
					this.labH102.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button34.BackColor = Color.DarkSlateBlue;
					this.labH100.BackColor = Color.DarkSlateBlue;
					this.labH101.BackColor = Color.DarkSlateBlue;
					this.labH102.BackColor = Color.DarkSlateBlue;
					this.labH100.Text = "";
					this.labH101.Text = "";
					this.labH102.Font = new Font(this.labH102.Font.FontFamily, 11);
					this.labH102.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button34.BackColor = Color.Turquoise;
					this.labH100.BackColor = Color.Turquoise;
					this.labH101.BackColor = Color.Turquoise;
					this.labH102.BackColor = Color.Turquoise;
					this.button34.ForeColor = Color.Black;
					this.labH100.ForeColor = Color.Black;
					this.labH101.ForeColor = Color.Black;
					this.labH102.ForeColor = Color.Black;
					this.labH100.Text = "";
					this.labH101.Text = "";
					this.labH102.Font = new Font(this.labH102.Font.FontFamily, 10);
					this.labH102.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button34.BackColor = Color.Navy;
					this.labH100.BackColor = Color.Navy;
					this.labH101.BackColor = Color.Navy;
					this.labH102.BackColor = Color.Navy;
					this.button34.ForeColor = Color.White;
					this.labH100.ForeColor = Color.White;
					this.labH101.ForeColor = Color.White;
					this.labH102.ForeColor = Color.White;
					this.labH100.Text = "";
					this.labH101.Text = "";
					this.labH102.Font = new Font(this.labH102.Font.FontFamily, 10);
					this.labH102.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button34.BackColor = Color.OrangeRed;
					this.labH100.BackColor = Color.OrangeRed;
					this.labH101.BackColor = Color.OrangeRed;
					this.labH102.BackColor = Color.OrangeRed;
					this.button34.ForeColor = Color.White;
					this.labH100.ForeColor = Color.White;
					this.labH101.ForeColor = Color.White;
					this.labH102.ForeColor = Color.White;
					this.labH100.Text = "";
					this.labH101.Text = "";
					this.labH102.Font = new Font(this.labH102.Font.FontFamily, 10);
					this.labH102.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button34.BackColor = Color.Orange;
					this.labH100.BackColor = Color.Orange;
					this.labH101.BackColor = Color.Orange;
					this.labH102.BackColor = Color.Orange;
					this.button34.ForeColor = Color.White;
					this.labH100.ForeColor = Color.White;
					this.labH101.ForeColor = Color.White;
					this.labH102.ForeColor = Color.White;
					this.labH100.Text = "";
					this.labH101.Text = "";
					this.labH102.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH102.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h34
			#region h35
			if (NurmdHabitacion==35)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button35.BackColor = Color.SteelBlue;
						this.labH103.BackColor = Color.SteelBlue;
						this.labH104.BackColor = Color.SteelBlue;
						this.labH105.BackColor = Color.SteelBlue;
						this.labH103.Text = Hinicia;
						this.labH104.Text = HFinaliza;
						this.labH105.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button35.BackColor = Color.Firebrick;
						this.labH103.BackColor = Color.Firebrick;
						this.labH104.BackColor = Color.Firebrick;
						this.labH105.BackColor = Color.Firebrick;
						this.labH103.Text = Hinicia;
						this.labH104.Text = HFinaliza;
						this.labH105.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button35.BackColor = Color.Orange;
							this.labH103.BackColor = Color.Orange;
							this.labH104.BackColor = Color.Orange;
							this.labH105.BackColor = Color.Orange;
							this.labH103.Text = Hinicia;
							this.labH104.Text = HFinaliza;
							this.labH105.Text = "T. QUEDA";
						}
						else
						{
							this.button35.BackColor = Color.Red;
							this.labH103.BackColor = Color.Red;
							this.labH104.BackColor = Color.Red;
							this.labH105.BackColor = Color.Red;
							this.labH103.Text = Hinicia;
							this.labH104.Text = HFinaliza;
							this.labH105.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button35.BackColor = Color.Yellow;
						this.button35.ForeColor = Color.Black;
						this.labH103.BackColor = Color.Yellow;
						this.labH104.BackColor = Color.Yellow;
						this.labH105.BackColor = Color.Yellow;
						this.labH103.ForeColor = Color.Black;
						this.labH104.ForeColor = Color.Black;
						this.labH105.ForeColor = Color.Black;
						this.labH103.Text = Hinicia;
						this.labH104.Text = HFinaliza;
						this.labH105.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button35.BackColor = Color.Silver;
					this.labH103.BackColor = Color.Silver;
					this.labH104.BackColor = Color.Silver;
					this.labH105.BackColor = Color.Silver;
					this.labH105.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button35.BackColor = Color.SaddleBrown;
						this.labH103.BackColor = Color.SaddleBrown;
						this.labH104.BackColor = Color.SaddleBrown;
						this.labH105.BackColor = Color.SaddleBrown;
						this.labH103.Text = hri;
						this.labH104.Text = hrf;
						this.labH105.Text = "RESERVADO";
					}
					else
					{
						this.button35.BackColor = Color.Green;
						this.labH103.BackColor = Color.Green;
						this.labH104.BackColor = Color.Green;
						this.labH105.BackColor = Color.Green;
						this.labH103.Text = "00:00";
						this.labH104.Text = "00:00";
						this.labH105.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn35.Visible = true;
					this.labelEn35.BackColor = Color.Red;
					this.labelEn35.ForeColor = Color.White;
					this.labelEn35.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn35.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button35.BackColor = Color.MediumPurple;
					this.labH103.BackColor = Color.MediumPurple;
					this.labH104.BackColor = Color.MediumPurple;
					this.labH105.BackColor = Color.MediumPurple;
					this.labH103.Text = "";
					this.labH104.Text = "";
					this.labH105.Font = new Font(this.labH105.Font.FontFamily, 10);
					this.labH105.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button35.BackColor = Color.DarkSlateBlue;
					this.labH103.BackColor = Color.DarkSlateBlue;
					this.labH104.BackColor = Color.DarkSlateBlue;
					this.labH105.BackColor = Color.DarkSlateBlue;
					this.labH103.Text = "";
					this.labH104.Text = "";
					this.labH105.Font = new Font(this.labH105.Font.FontFamily, 11);
					this.labH105.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button35.BackColor = Color.Turquoise;
					this.labH103.BackColor = Color.Turquoise;
					this.labH104.BackColor = Color.Turquoise;
					this.labH105.BackColor = Color.Turquoise;
					this.button35.ForeColor = Color.Black;
					this.labH103.ForeColor = Color.Black;
					this.labH104.ForeColor = Color.Black;
					this.labH105.ForeColor = Color.Black;
					this.labH103.Text = "";
					this.labH104.Text = "";
					this.labH105.Font = new Font(this.labH105.Font.FontFamily, 10);
					this.labH105.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button35.BackColor = Color.Navy;
					this.labH103.BackColor = Color.Navy;
					this.labH104.BackColor = Color.Navy;
					this.labH105.BackColor = Color.Navy;
					this.button35.ForeColor = Color.White;
					this.labH103.ForeColor = Color.White;
					this.labH104.ForeColor = Color.White;
					this.labH105.ForeColor = Color.White;
					this.labH103.Text = "";
					this.labH104.Text = "";
					this.labH105.Font = new Font(this.labH105.Font.FontFamily, 10);
					this.labH105.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button35.BackColor = Color.OrangeRed;
					this.labH103.BackColor = Color.OrangeRed;
					this.labH104.BackColor = Color.OrangeRed;
					this.labH105.BackColor = Color.OrangeRed;
					this.button35.ForeColor = Color.White;
					this.labH103.ForeColor = Color.White;
					this.labH104.ForeColor = Color.White;
					this.labH105.ForeColor = Color.White;
					this.labH103.Text = "";
					this.labH104.Text = "";
					this.labH105.Font = new Font(this.labH105.Font.FontFamily, 10);
					this.labH105.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button35.BackColor = Color.Orange;
					this.labH103.BackColor = Color.Orange;
					this.labH104.BackColor = Color.Orange;
					this.labH105.BackColor = Color.Orange;
					this.button35.ForeColor = Color.White;
					this.labH103.ForeColor = Color.White;
					this.labH104.ForeColor = Color.White;
					this.labH105.ForeColor = Color.White;
					this.labH103.Text = "";
					this.labH104.Text = "";
					this.labH105.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH105.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h35
			#region h36
			if (NurmdHabitacion==36)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button36.BackColor = Color.SteelBlue;
						this.labH106.BackColor = Color.SteelBlue;
						this.labH107.BackColor = Color.SteelBlue;
						this.labH108.BackColor = Color.SteelBlue;
						this.labH106.Text = Hinicia;
						this.labH107.Text = HFinaliza;
						this.labH108.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button36.BackColor = Color.Firebrick;
						this.labH106.BackColor = Color.Firebrick;
						this.labH107.BackColor = Color.Firebrick;
						this.labH108.BackColor = Color.Firebrick;
						this.labH106.Text = Hinicia;
						this.labH107.Text = HFinaliza;
						this.labH108.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button36.BackColor = Color.Orange;
							this.labH106.BackColor = Color.Orange;
							this.labH107.BackColor = Color.Orange;
							this.labH108.BackColor = Color.Orange;
							this.labH106.Text = Hinicia;
							this.labH107.Text = HFinaliza;
							this.labH108.Text = "T. QUEDA";
						}
						else
						{
							this.button36.BackColor = Color.Red;
							this.labH106.BackColor = Color.Red;
							this.labH107.BackColor = Color.Red;
							this.labH108.BackColor = Color.Red;
							this.labH106.Text = Hinicia;
							this.labH107.Text = HFinaliza;
							this.labH108.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button36.BackColor = Color.Yellow;
						this.button36.ForeColor = Color.Black;
						this.labH106.BackColor = Color.Yellow;
						this.labH107.BackColor = Color.Yellow;
						this.labH108.BackColor = Color.Yellow;
						this.labH106.ForeColor = Color.Black;
						this.labH107.ForeColor = Color.Black;
						this.labH108.ForeColor = Color.Black;
						this.labH106.Text = Hinicia;
						this.labH107.Text = HFinaliza;
						this.labH108.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button36.BackColor = Color.Silver;
					this.labH106.BackColor = Color.Silver;
					this.labH107.BackColor = Color.Silver;
					this.labH108.BackColor = Color.Silver;
					this.labH108.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button36.BackColor = Color.SaddleBrown;
						this.labH106.BackColor = Color.SaddleBrown;
						this.labH107.BackColor = Color.SaddleBrown;
						this.labH108.BackColor = Color.SaddleBrown;
						this.labH106.Text = hri;
						this.labH107.Text = hrf;
						this.labH108.Text = "RESERVADO";
					}
					else
					{
						this.button36.BackColor = Color.Green;
						this.labH106.BackColor = Color.Green;
						this.labH107.BackColor = Color.Green;
						this.labH108.BackColor = Color.Green;
						this.labH106.Text = "00:00";
						this.labH107.Text = "00:00";
						this.labH108.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn36.Visible = true;
					this.labelEn36.BackColor = Color.Red;
					this.labelEn36.ForeColor = Color.White;
					this.labelEn36.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn36.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button36.BackColor = Color.MediumPurple;
					this.labH106.BackColor = Color.MediumPurple;
					this.labH107.BackColor = Color.MediumPurple;
					this.labH108.BackColor = Color.MediumPurple;
					this.labH106.Text = "";
					this.labH107.Text = "";
					this.labH108.Font = new Font(this.labH108.Font.FontFamily, 10);
					this.labH108.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button36.BackColor = Color.DarkSlateBlue;
					this.labH106.BackColor = Color.DarkSlateBlue;
					this.labH107.BackColor = Color.DarkSlateBlue;
					this.labH108.BackColor = Color.DarkSlateBlue;
					this.labH106.Text = "";
					this.labH107.Text = "";
					this.labH108.Font = new Font(this.labH108.Font.FontFamily, 11);
					this.labH108.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button36.BackColor = Color.Turquoise;
					this.labH106.BackColor = Color.Turquoise;
					this.labH107.BackColor = Color.Turquoise;
					this.labH108.BackColor = Color.Turquoise;
					this.button36.ForeColor = Color.Black;
					this.labH106.ForeColor = Color.Black;
					this.labH107.ForeColor = Color.Black;
					this.labH108.ForeColor = Color.Black;
					this.labH106.Text = "";
					this.labH107.Text = "";
					this.labH108.Font = new Font(this.labH108.Font.FontFamily, 10);
					this.labH108.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button36.BackColor = Color.Navy;
					this.labH106.BackColor = Color.Navy;
					this.labH107.BackColor = Color.Navy;
					this.labH108.BackColor = Color.Navy;
					this.button36.ForeColor = Color.White;
					this.labH106.ForeColor = Color.White;
					this.labH107.ForeColor = Color.White;
					this.labH108.ForeColor = Color.White;
					this.labH106.Text = "";
					this.labH107.Text = "";
					this.labH108.Font = new Font(this.labH108.Font.FontFamily, 10);
					this.labH108.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button36.BackColor = Color.OrangeRed;
					this.labH106.BackColor = Color.OrangeRed;
					this.labH107.BackColor = Color.OrangeRed;
					this.labH108.BackColor = Color.OrangeRed;
					this.button36.ForeColor = Color.White;
					this.labH106.ForeColor = Color.White;
					this.labH107.ForeColor = Color.White;
					this.labH108.ForeColor = Color.White;
					this.labH106.Text = "";
					this.labH107.Text = "";
					this.labH108.Font = new Font(this.labH108.Font.FontFamily, 10);
					this.labH108.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button36.BackColor = Color.Orange;
					this.labH106.BackColor = Color.Orange;
					this.labH107.BackColor = Color.Orange;
					this.labH108.BackColor = Color.Orange;
					this.button36.ForeColor = Color.White;
					this.labH106.ForeColor = Color.White;
					this.labH107.ForeColor = Color.White;
					this.labH108.ForeColor = Color.White;
					this.labH106.Text = "";
					this.labH107.Text = "";
					this.labH108.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH108.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h36
			#region h37
			if (NurmdHabitacion==37)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button37.BackColor = Color.SteelBlue;
						this.labH109.BackColor = Color.SteelBlue;
						this.labH110.BackColor = Color.SteelBlue;
						this.labH111.BackColor = Color.SteelBlue;
						this.labH109.Text = Hinicia;
						this.labH110.Text = HFinaliza;
						this.labH111.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button37.BackColor = Color.Firebrick;
						this.labH109.BackColor = Color.Firebrick;
						this.labH110.BackColor = Color.Firebrick;
						this.labH111.BackColor = Color.Firebrick;
						this.labH109.Text = Hinicia;
						this.labH110.Text = HFinaliza;
						this.labH111.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button37.BackColor = Color.Orange;
							this.labH109.BackColor = Color.Orange;
							this.labH110.BackColor = Color.Orange;
							this.labH111.BackColor = Color.Orange;
							this.labH109.Text = Hinicia;
							this.labH110.Text = HFinaliza;
							this.labH111.Text = "T. QUEDA";
						}
						else
						{
							this.button37.BackColor = Color.Red;
							this.labH109.BackColor = Color.Red;
							this.labH110.BackColor = Color.Red;
							this.labH111.BackColor = Color.Red;
							this.labH109.Text = Hinicia;
							this.labH110.Text = HFinaliza;
							this.labH111.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button37.BackColor = Color.Yellow;
						this.button37.ForeColor = Color.Black;
						this.labH109.BackColor = Color.Yellow;
						this.labH110.BackColor = Color.Yellow;
						this.labH111.BackColor = Color.Yellow;
						this.labH109.ForeColor = Color.Black;
						this.labH110.ForeColor = Color.Black;
						this.labH111.ForeColor = Color.Black;
						this.labH109.Text = Hinicia;
						this.labH110.Text = HFinaliza;
						this.labH111.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button37.BackColor = Color.Silver;
					this.labH109.BackColor = Color.Silver;
					this.labH110.BackColor = Color.Silver;
					this.labH111.BackColor = Color.Silver;
					this.labH111.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button37.BackColor = Color.SaddleBrown;
						this.labH109.BackColor = Color.SaddleBrown;
						this.labH110.BackColor = Color.SaddleBrown;
						this.labH111.BackColor = Color.SaddleBrown;
						this.labH109.Text = hri;
						this.labH110.Text = hrf;
						this.labH111.Text = "RESERVADO";
					}
					else
					{
						this.button37.BackColor = Color.Green;
						this.labH109.BackColor = Color.Green;
						this.labH110.BackColor = Color.Green;
						this.labH111.BackColor = Color.Green;
						this.labH109.Text = "00:00";
						this.labH110.Text = "00:00";
						this.labH111.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn37.Visible = true;
					this.labelEn37.BackColor = Color.Red;
					this.labelEn37.ForeColor = Color.White;
					this.labelEn37.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn37.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button37.BackColor = Color.MediumPurple;
					this.labH109.BackColor = Color.MediumPurple;
					this.labH110.BackColor = Color.MediumPurple;
					this.labH111.BackColor = Color.MediumPurple;
					this.labH109.Text = "";
					this.labH110.Text = "";
					this.labH111.Font = new Font(this.labH111.Font.FontFamily, 10);
					this.labH111.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button37.BackColor = Color.DarkSlateBlue;
					this.labH109.BackColor = Color.DarkSlateBlue;
					this.labH110.BackColor = Color.DarkSlateBlue;
					this.labH111.BackColor = Color.DarkSlateBlue;
					this.labH109.Text = "";
					this.labH110.Text = "";
					this.labH111.Font = new Font(this.labH111.Font.FontFamily, 11);
					this.labH111.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button37.BackColor = Color.Turquoise;
					this.labH109.BackColor = Color.Turquoise;
					this.labH110.BackColor = Color.Turquoise;
					this.labH111.BackColor = Color.Turquoise;
					this.button37.ForeColor = Color.Black;
					this.labH109.ForeColor = Color.Black;
					this.labH110.ForeColor = Color.Black;
					this.labH111.ForeColor = Color.Black;
					this.labH109.Text = "";
					this.labH110.Text = "";
					this.labH111.Font = new Font(this.labH111.Font.FontFamily, 10);
					this.labH111.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button37.BackColor = Color.Navy;
					this.labH109.BackColor = Color.Navy;
					this.labH110.BackColor = Color.Navy;
					this.labH111.BackColor = Color.Navy;
					this.button37.ForeColor = Color.White;
					this.labH109.ForeColor = Color.White;
					this.labH110.ForeColor = Color.White;
					this.labH111.ForeColor = Color.White;
					this.labH109.Text = "";
					this.labH110.Text = "";
					this.labH111.Font = new Font(this.labH111.Font.FontFamily, 10);
					this.labH111.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button37.BackColor = Color.OrangeRed;
					this.labH109.BackColor = Color.OrangeRed;
					this.labH110.BackColor = Color.OrangeRed;
					this.labH111.BackColor = Color.OrangeRed;
					this.button37.ForeColor = Color.White;
					this.labH109.ForeColor = Color.White;
					this.labH110.ForeColor = Color.White;
					this.labH111.ForeColor = Color.White;
					this.labH109.Text = "";
					this.labH110.Text = "";
					this.labH111.Font = new Font(this.labH111.Font.FontFamily, 10);
					this.labH111.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button37.BackColor = Color.Orange;
					this.labH109.BackColor = Color.Orange;
					this.labH110.BackColor = Color.Orange;
					this.labH111.BackColor = Color.Orange;
					this.button37.ForeColor = Color.White;
					this.labH109.ForeColor = Color.White;
					this.labH110.ForeColor = Color.White;
					this.labH111.ForeColor = Color.White;
					this.labH109.Text = "";
					this.labH110.Text = "";
					this.labH111.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH111.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h37
			#region h38
			if (NurmdHabitacion==38)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button38.BackColor = Color.SteelBlue;
						this.labH112.BackColor = Color.SteelBlue;
						this.labH113.BackColor = Color.SteelBlue;
						this.labH114.BackColor = Color.SteelBlue;
						this.labH112.Text = Hinicia;
						this.labH113.Text = HFinaliza;
						this.labH114.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button38.BackColor = Color.Firebrick;
						this.labH112.BackColor = Color.Firebrick;
						this.labH113.BackColor = Color.Firebrick;
						this.labH114.BackColor = Color.Firebrick;
						this.labH112.Text = Hinicia;
						this.labH113.Text = HFinaliza;
						this.labH114.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button38.BackColor = Color.Orange;
							this.labH112.BackColor = Color.Orange;
							this.labH113.BackColor = Color.Orange;
							this.labH114.BackColor = Color.Orange;
							this.labH112.Text = Hinicia;
							this.labH113.Text = HFinaliza;
							this.labH114.Text = "T. QUEDA";
						}
						else
						{
							this.button38.BackColor = Color.Red;
							this.labH112.BackColor = Color.Red;
							this.labH113.BackColor = Color.Red;
							this.labH114.BackColor = Color.Red;
							this.labH112.Text = Hinicia;
							this.labH113.Text = HFinaliza;
							this.labH114.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button38.BackColor = Color.Yellow;
						this.button38.ForeColor = Color.Black;
						this.labH112.BackColor = Color.Yellow;
						this.labH113.BackColor = Color.Yellow;
						this.labH114.BackColor = Color.Yellow;
						this.labH112.ForeColor = Color.Black;
						this.labH113.ForeColor = Color.Black;
						this.labH114.ForeColor = Color.Black;
						this.labH112.Text = Hinicia;
						this.labH113.Text = HFinaliza;
						this.labH114.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button38.BackColor = Color.Silver;
					this.labH112.BackColor = Color.Silver;
					this.labH113.BackColor = Color.Silver;
					this.labH114.BackColor = Color.Silver;
					this.labH114.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button38.BackColor = Color.SaddleBrown;
						this.labH112.BackColor = Color.SaddleBrown;
						this.labH113.BackColor = Color.SaddleBrown;
						this.labH114.BackColor = Color.SaddleBrown;
						this.labH112.Text = hri;
						this.labH113.Text = hrf;
						this.labH114.Text = "RESERVADO";
					}
					else
					{
						this.button38.BackColor = Color.Green;
						this.labH112.BackColor = Color.Green;
						this.labH113.BackColor = Color.Green;
						this.labH114.BackColor = Color.Green;
						this.labH112.Text = "00:00";
						this.labH113.Text = "00:00";
						this.labH114.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn38.Visible = true;
					this.labelEn38.BackColor = Color.Red;
					this.labelEn38.ForeColor = Color.White;
					this.labelEn38.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn38.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button38.BackColor = Color.MediumPurple;
					this.labH112.BackColor = Color.MediumPurple;
					this.labH113.BackColor = Color.MediumPurple;
					this.labH114.BackColor = Color.MediumPurple;
					this.labH112.Text = "";
					this.labH113.Text = "";
					this.labH114.Font = new Font(this.labH114.Font.FontFamily, 10);
					this.labH114.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button38.BackColor = Color.DarkSlateBlue;
					this.labH112.BackColor = Color.DarkSlateBlue;
					this.labH113.BackColor = Color.DarkSlateBlue;
					this.labH114.BackColor = Color.DarkSlateBlue;
					this.labH112.Text = "";
					this.labH113.Text = "";
					this.labH114.Font = new Font(this.labH114.Font.FontFamily, 11);
					this.labH114.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button38.BackColor = Color.Turquoise;
					this.labH112.BackColor = Color.Turquoise;
					this.labH113.BackColor = Color.Turquoise;
					this.labH114.BackColor = Color.Turquoise;
					this.button38.ForeColor = Color.Black;
					this.labH112.ForeColor = Color.Black;
					this.labH113.ForeColor = Color.Black;
					this.labH114.ForeColor = Color.Black;
					this.labH112.Text = "";
					this.labH113.Text = "";
					this.labH114.Font = new Font(this.labH114.Font.FontFamily, 10);
					this.labH114.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button38.BackColor = Color.Navy;
					this.labH112.BackColor = Color.Navy;
					this.labH113.BackColor = Color.Navy;
					this.labH114.BackColor = Color.Navy;
					this.button38.ForeColor = Color.White;
					this.labH112.ForeColor = Color.White;
					this.labH113.ForeColor = Color.White;
					this.labH114.ForeColor = Color.White;
					this.labH112.Text = "";
					this.labH113.Text = "";
					this.labH114.Font = new Font(this.labH114.Font.FontFamily, 10);
					this.labH114.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button38.BackColor = Color.OrangeRed;
					this.labH112.BackColor = Color.OrangeRed;
					this.labH113.BackColor = Color.OrangeRed;
					this.labH114.BackColor = Color.OrangeRed;
					this.button38.ForeColor = Color.White;
					this.labH112.ForeColor = Color.White;
					this.labH113.ForeColor = Color.White;
					this.labH114.ForeColor = Color.White;
					this.labH112.Text = "";
					this.labH113.Text = "";
					this.labH114.Font = new Font(this.labH114.Font.FontFamily, 10);
					this.labH114.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button38.BackColor = Color.Orange;
					this.labH112.BackColor = Color.Orange;
					this.labH113.BackColor = Color.Orange;
					this.labH114.BackColor = Color.Orange;
					this.button38.ForeColor = Color.White;
					this.labH112.ForeColor = Color.White;
					this.labH113.ForeColor = Color.White;
					this.labH114.ForeColor = Color.White;
					this.labH112.Text = "";
					this.labH113.Text = "";
					this.labH114.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH114.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h38
			#region h39
			if (NurmdHabitacion==39)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button39.BackColor = Color.SteelBlue;
						this.labH115.BackColor = Color.SteelBlue;
						this.labH116.BackColor = Color.SteelBlue;
						this.labH117.BackColor = Color.SteelBlue;
						this.labH115.Text = Hinicia;
						this.labH116.Text = HFinaliza;
						this.labH117.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button39.BackColor = Color.Firebrick;
						this.labH115.BackColor = Color.Firebrick;
						this.labH116.BackColor = Color.Firebrick;
						this.labH117.BackColor = Color.Firebrick;
						this.labH115.Text = Hinicia;
						this.labH116.Text = HFinaliza;
						this.labH117.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button39.BackColor = Color.Orange;
							this.labH115.BackColor = Color.Orange;
							this.labH116.BackColor = Color.Orange;
							this.labH117.BackColor = Color.Orange;
							this.labH115.Text = Hinicia;
							this.labH116.Text = HFinaliza;
							this.labH117.Text = "T. QUEDA";
						}
						else
						{
							this.button39.BackColor = Color.Red;
							this.labH115.BackColor = Color.Red;
							this.labH116.BackColor = Color.Red;
							this.labH117.BackColor = Color.Red;
							this.labH115.Text = Hinicia;
							this.labH116.Text = HFinaliza;
							this.labH117.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button39.BackColor = Color.Yellow;
						this.button39.ForeColor = Color.Black;
						this.labH115.BackColor = Color.Yellow;
						this.labH116.BackColor = Color.Yellow;
						this.labH117.BackColor = Color.Yellow;
						this.labH115.ForeColor = Color.Black;
						this.labH116.ForeColor = Color.Black;
						this.labH117.ForeColor = Color.Black;
						this.labH115.Text = Hinicia;
						this.labH116.Text = HFinaliza;
						this.labH117.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button39.BackColor = Color.Silver;
					this.labH115.BackColor = Color.Silver;
					this.labH116.BackColor = Color.Silver;
					this.labH117.BackColor = Color.Silver;
					this.labH117.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button39.BackColor = Color.SaddleBrown;
						this.labH115.BackColor = Color.SaddleBrown;
						this.labH116.BackColor = Color.SaddleBrown;
						this.labH117.BackColor = Color.SaddleBrown;
						this.labH115.Text = hri;
						this.labH116.Text = hrf;
						this.labH117.Text = "RESERVADO";
					}
					else
					{
						this.button39.BackColor = Color.Green;
						this.labH115.BackColor = Color.Green;
						this.labH116.BackColor = Color.Green;
						this.labH117.BackColor = Color.Green;
						this.labH115.Text = "00:00";
						this.labH116.Text = "00:00";
						this.labH117.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn39.Visible = true;
					this.labelEn39.BackColor = Color.Red;
					this.labelEn39.ForeColor = Color.White;
					this.labelEn39.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn39.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button39.BackColor = Color.MediumPurple;
					this.labH115.BackColor = Color.MediumPurple;
					this.labH116.BackColor = Color.MediumPurple;
					this.labH117.BackColor = Color.MediumPurple;
					this.labH115.Text = "";
					this.labH116.Text = "";
					this.labH117.Font = new Font(this.labH117.Font.FontFamily, 10);
					this.labH117.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button39.BackColor = Color.DarkSlateBlue;
					this.labH115.BackColor = Color.DarkSlateBlue;
					this.labH116.BackColor = Color.DarkSlateBlue;
					this.labH117.BackColor = Color.DarkSlateBlue;
					this.labH115.Text = "";
					this.labH116.Text = "";
					this.labH117.Font = new Font(this.labH117.Font.FontFamily, 11);
					this.labH117.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button39.BackColor = Color.Turquoise;
					this.labH115.BackColor = Color.Turquoise;
					this.labH116.BackColor = Color.Turquoise;
					this.labH117.BackColor = Color.Turquoise;
					this.button39.ForeColor = Color.Black;
					this.labH115.ForeColor = Color.Black;
					this.labH116.ForeColor = Color.Black;
					this.labH117.ForeColor = Color.Black;
					this.labH115.Text = "";
					this.labH116.Text = "";
					this.labH117.Font = new Font(this.labH117.Font.FontFamily, 10);
					this.labH117.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button39.BackColor = Color.Navy;
					this.labH115.BackColor = Color.Navy;
					this.labH116.BackColor = Color.Navy;
					this.labH117.BackColor = Color.Navy;
					this.button39.ForeColor = Color.White;
					this.labH115.ForeColor = Color.White;
					this.labH116.ForeColor = Color.White;
					this.labH117.ForeColor = Color.White;
					this.labH115.Text = "";
					this.labH116.Text = "";
					this.labH117.Font = new Font(this.labH117.Font.FontFamily, 10);
					this.labH117.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button39.BackColor = Color.OrangeRed;
					this.labH115.BackColor = Color.OrangeRed;
					this.labH116.BackColor = Color.OrangeRed;
					this.labH117.BackColor = Color.OrangeRed;
					this.button39.ForeColor = Color.White;
					this.labH115.ForeColor = Color.White;
					this.labH116.ForeColor = Color.White;
					this.labH117.ForeColor = Color.White;
					this.labH115.Text = "";
					this.labH116.Text = "";
					this.labH117.Font = new Font(this.labH117.Font.FontFamily, 10);
					this.labH117.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button39.BackColor = Color.Orange;
					this.labH115.BackColor = Color.Orange;
					this.labH116.BackColor = Color.Orange;
					this.labH117.BackColor = Color.Orange;
					this.button39.ForeColor = Color.White;
					this.labH115.ForeColor = Color.White;
					this.labH116.ForeColor = Color.White;
					this.labH117.ForeColor = Color.White;
					this.labH115.Text = "";
					this.labH116.Text = "";
					this.labH117.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH117.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h39
			#region h40
			if (NurmdHabitacion==40)
			{
				if (e_hab==2)
				{
					if (dif_mins>=0 && dif_mins <=10) // PASADO HASTA 10 MINUTOS
					{
						this.button40.BackColor = Color.SteelBlue;
						this.labH118.BackColor = Color.SteelBlue;
						this.labH119.BackColor = Color.SteelBlue;
						this.labH120.BackColor = Color.SteelBlue;
						this.labH118.Text = Hinicia;
						this.labH119.Text = HFinaliza;
						this.labH120.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else if (dif_mins>=-10 && dif_mins<0) // 10 MINUTOS ANTES QUE SE CUMPLA EL TIEMPO
					{							
						this.button40.BackColor = Color.Firebrick;
						this.labH118.BackColor = Color.Firebrick;
						this.labH119.BackColor = Color.Firebrick;
						this.labH120.BackColor = Color.Firebrick;
						this.labH118.Text = Hinicia;
						this.labH119.Text = HFinaliza;
						this.labH120.Text = "FALTAN:"+dif_mins.ToString()+" Min.";						
					}
					else // OCUPADA
					{
						if(bToqueQueda)
						{
							this.button40.BackColor = Color.Orange;
							this.labH118.BackColor = Color.Orange;
							this.labH119.BackColor = Color.Orange;
							this.labH120.BackColor = Color.Orange;
							this.labH118.Text = Hinicia;
							this.labH119.Text = HFinaliza;
							this.labH120.Text = "T. QUEDA";
						}
						else
						{
							this.button40.BackColor = Color.Red;
							this.labH118.BackColor = Color.Red;
							this.labH119.BackColor = Color.Red;
							this.labH120.BackColor = Color.Red;
							this.labH118.Text = Hinicia;
							this.labH119.Text = HFinaliza;
							this.labH120.Text = "OCUPADO";
						}
					}
				}
				if (e_hab==3)
				{

					{
						this.button40.BackColor = Color.Yellow;
						this.button40.ForeColor = Color.Black;
						this.labH118.BackColor = Color.Yellow;
						this.labH119.BackColor = Color.Yellow;
						this.labH120.BackColor = Color.Yellow;
						this.labH118.ForeColor = Color.Black;
						this.labH119.ForeColor = Color.Black;
						this.labH120.ForeColor = Color.Black;
						this.labH118.Text = Hinicia;
						this.labH119.Text = HFinaliza;
						this.labH120.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button40.BackColor = Color.Silver;
					this.labH118.BackColor = Color.Silver;
					this.labH119.BackColor = Color.Silver;
					this.labH120.BackColor = Color.Silver;
					this.labH120.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button40.BackColor = Color.SaddleBrown;
						this.labH118.BackColor = Color.SaddleBrown;
						this.labH119.BackColor = Color.SaddleBrown;
						this.labH120.BackColor = Color.SaddleBrown;
						this.labH118.Text = hri;
						this.labH119.Text = hrf;
						this.labH120.Text = "RESERVADO";
					}
					else
					{
						this.button40.BackColor = Color.Green;
						this.labH118.BackColor = Color.Green;
						this.labH119.BackColor = Color.Green;
						this.labH120.BackColor = Color.Green;
						this.labH118.Text = "00:00";
						this.labH119.Text = "00:00";
						this.labH120.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn40.Visible = true;
					this.labelEn40.BackColor = Color.Red;
					this.labelEn40.ForeColor = Color.White;
					this.labelEn40.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn40.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button40.BackColor = Color.MediumPurple;
					this.labH118.BackColor = Color.MediumPurple;
					this.labH119.BackColor = Color.MediumPurple;
					this.labH120.BackColor = Color.MediumPurple;
					this.labH118.Text = "";
					this.labH119.Text = "";
					this.labH120.Font = new Font(this.labH120.Font.FontFamily, 10);
					this.labH120.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button40.BackColor = Color.DarkSlateBlue;
					this.labH118.BackColor = Color.DarkSlateBlue;
					this.labH119.BackColor = Color.DarkSlateBlue;
					this.labH120.BackColor = Color.DarkSlateBlue;
					this.labH118.Text = "";
					this.labH119.Text = "";
					this.labH120.Font = new Font(this.labH120.Font.FontFamily, 11);
					this.labH120.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button40.BackColor = Color.Turquoise;
					this.labH118.BackColor = Color.Turquoise;
					this.labH119.BackColor = Color.Turquoise;
					this.labH120.BackColor = Color.Turquoise;
					this.button40.ForeColor = Color.Black;
					this.labH118.ForeColor = Color.Black;
					this.labH119.ForeColor = Color.Black;
					this.labH120.ForeColor = Color.Black;
					this.labH118.Text = "";
					this.labH119.Text = "";
					this.labH120.Font = new Font(this.labH120.Font.FontFamily, 10);
					this.labH120.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button40.BackColor = Color.Navy;
					this.labH118.BackColor = Color.Navy;
					this.labH119.BackColor = Color.Navy;
					this.labH120.BackColor = Color.Navy;
					this.button40.ForeColor = Color.White;
					this.labH118.ForeColor = Color.White;
					this.labH119.ForeColor = Color.White;
					this.labH120.ForeColor = Color.White;
					this.labH118.Text = "";
					this.labH119.Text = "";
					this.labH120.Font = new Font(this.labH120.Font.FontFamily, 10);
					this.labH120.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button40.BackColor = Color.OrangeRed;
					this.labH118.BackColor = Color.OrangeRed;
					this.labH119.BackColor = Color.OrangeRed;
					this.labH120.BackColor = Color.OrangeRed;
					this.button40.ForeColor = Color.White;
					this.labH118.ForeColor = Color.White;
					this.labH119.ForeColor = Color.White;
					this.labH120.ForeColor = Color.White;
					this.labH118.Text = "";
					this.labH119.Text = "";
					this.labH120.Font = new Font(this.labH120.Font.FontFamily, 10);
					this.labH120.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button40.BackColor = Color.Orange;
					this.labH118.BackColor = Color.Orange;
					this.labH119.BackColor = Color.Orange;
					this.labH120.BackColor = Color.Orange;
					this.button40.ForeColor = Color.White;
					this.labH118.ForeColor = Color.White;
					this.labH119.ForeColor = Color.White;
					this.labH120.ForeColor = Color.White;
					this.labH118.Text = "";
					this.labH119.Text = "";
					this.labH120.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH120.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h40
			#region h41
			if (NurmdHabitacion==41)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button41.BackColor = Color.SteelBlue;
						this.labH121.BackColor = Color.SteelBlue;
						this.labH122.BackColor = Color.SteelBlue;
						this.labH123.BackColor = Color.SteelBlue;
						this.labH121.Text = Hinicia;
						this.labH122.Text = HFinaliza;
						this.labH123.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else 
					{
						this.button41.BackColor = Color.Red;
						this.labH121.BackColor = Color.Red;
						this.labH122.BackColor = Color.Red;
						this.labH123.BackColor = Color.Red;
						this.labH121.Text = Hinicia;
						this.labH122.Text = HFinaliza;
						this.labH123.Text = "OCUPADO";
					}
				}
				if (e_hab==3)
				{
	
					{
						this.button41.BackColor = Color.Yellow;
						this.button41.ForeColor = Color.Black;
						this.labH121.BackColor = Color.Yellow;
						this.labH122.BackColor = Color.Yellow;
						this.labH123.BackColor = Color.Yellow;
						this.labH121.ForeColor = Color.Black;
						this.labH122.ForeColor = Color.Black;
						this.labH123.ForeColor = Color.Black;
						this.labH121.Text = Hinicia;
						this.labH122.Text = HFinaliza;
						this.labH123.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button41.BackColor = Color.Silver;
					this.labH121.BackColor = Color.Silver;
					this.labH122.BackColor = Color.Silver;
					this.labH123.BackColor = Color.Silver;
					this.labH123.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button41.BackColor = Color.SaddleBrown;
						this.labH121.BackColor = Color.SaddleBrown;
						this.labH122.BackColor = Color.SaddleBrown;
						this.labH123.BackColor = Color.SaddleBrown;
						this.labH121.Text = hri;
						this.labH122.Text = hrf;
						this.labH123.Text = "RESERVADO";
					}
					else
					{
						this.button41.BackColor = Color.Green;
						this.labH121.BackColor = Color.Green;
						this.labH122.BackColor = Color.Green;
						this.labH123.BackColor = Color.Green;
						this.labH121.Text = "00:00";
						this.labH122.Text = "00:00";
						this.labH123.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn41.Visible = true;
					this.labelEn41.BackColor = Color.Red;
					this.labelEn41.ForeColor = Color.White;
					this.labelEn41.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn41.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button41.BackColor = Color.MediumPurple;
					this.labH121.BackColor = Color.MediumPurple;
					this.labH122.BackColor = Color.MediumPurple;
					this.labH123.BackColor = Color.MediumPurple;
					this.labH121.Text = "";
					this.labH122.Text = "";
					this.labH123.Font = new Font(this.labH123.Font.FontFamily, 10);
					this.labH123.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button41.BackColor = Color.DarkSlateBlue;
					this.labH121.BackColor = Color.DarkSlateBlue;
					this.labH122.BackColor = Color.DarkSlateBlue;
					this.labH123.BackColor = Color.DarkSlateBlue;
					this.labH121.Text = "";
					this.labH122.Text = "";
					this.labH123.Font = new Font(this.labH123.Font.FontFamily, 11);
					this.labH123.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button41.BackColor = Color.Turquoise;
					this.labH121.BackColor = Color.Turquoise;
					this.labH122.BackColor = Color.Turquoise;
					this.labH123.BackColor = Color.Turquoise;
					this.button41.ForeColor = Color.Black;
					this.labH121.ForeColor = Color.Black;
					this.labH122.ForeColor = Color.Black;
					this.labH123.ForeColor = Color.Black;
					this.labH121.Text = "";
					this.labH122.Text = "";
					this.labH123.Font = new Font(this.labH123.Font.FontFamily, 10);
					this.labH123.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button41.BackColor = Color.Navy;
					this.labH121.BackColor = Color.Navy;
					this.labH122.BackColor = Color.Navy;
					this.labH123.BackColor = Color.Navy;
					this.button41.ForeColor = Color.White;
					this.labH121.ForeColor = Color.White;
					this.labH122.ForeColor = Color.White;
					this.labH123.ForeColor = Color.White;
					this.labH121.Text = "";
					this.labH122.Text = "";
					this.labH123.Font = new Font(this.labH123.Font.FontFamily, 10);
					this.labH123.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button41.BackColor = Color.OrangeRed;
					this.labH121.BackColor = Color.OrangeRed;
					this.labH122.BackColor = Color.OrangeRed;
					this.labH123.BackColor = Color.OrangeRed;
					this.button41.ForeColor = Color.White;
					this.labH121.ForeColor = Color.White;
					this.labH122.ForeColor = Color.White;
					this.labH123.ForeColor = Color.White;
					this.labH121.Text = "";
					this.labH122.Text = "";
					this.labH123.Font = new Font(this.labH123.Font.FontFamily, 10);
					this.labH123.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button41.BackColor = Color.Orange;
					this.labH121.BackColor = Color.Orange;
					this.labH122.BackColor = Color.Orange;
					this.labH123.BackColor = Color.Orange;
					this.button41.ForeColor = Color.White;
					this.labH121.ForeColor = Color.White;
					this.labH122.ForeColor = Color.White;
					this.labH123.ForeColor = Color.White;
					this.labH121.Text = "";
					this.labH122.Text = "";
					this.labH123.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH123.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h41
			#region h42
			if (NurmdHabitacion==42)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button42.BackColor = Color.SteelBlue;
						this.labH124.BackColor = Color.SteelBlue;
						this.labH125.BackColor = Color.SteelBlue;
						this.labH126.BackColor = Color.SteelBlue;
						this.labH124.Text = Hinicia;
						this.labH125.Text = HFinaliza;
						this.labH126.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else 
					{
						this.button42.BackColor = Color.Red;
						this.labH124.BackColor = Color.Red;
						this.labH125.BackColor = Color.Red;
						this.labH126.BackColor = Color.Red;
						this.labH124.Text = Hinicia;
						this.labH125.Text = HFinaliza;
						this.labH126.Text = "OCUPADO";
					}
				}
				if (e_hab==3)
				{

					{
						this.button42.BackColor = Color.Yellow;
						this.button42.ForeColor = Color.Black;
						this.labH124.BackColor = Color.Yellow;
						this.labH125.BackColor = Color.Yellow;
						this.labH126.BackColor = Color.Yellow;
						this.labH124.ForeColor = Color.Black;
						this.labH125.ForeColor = Color.Black;
						this.labH126.ForeColor = Color.Black;
						this.labH124.Text = Hinicia;
						this.labH125.Text = HFinaliza;
						this.labH126.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button42.BackColor = Color.Silver;
					this.labH124.BackColor = Color.Silver;
					this.labH125.BackColor = Color.Silver;
					this.labH126.BackColor = Color.Silver;
					this.labH126.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button42.BackColor = Color.SaddleBrown;
						this.labH124.BackColor = Color.SaddleBrown;
						this.labH125.BackColor = Color.SaddleBrown;
						this.labH126.BackColor = Color.SaddleBrown;
						this.labH124.Text = hri;
						this.labH125.Text = hrf;
						this.labH126.Text = "RESERVADO";
					}
					else
					{
						this.button42.BackColor = Color.Green;
						this.labH124.BackColor = Color.Green;
						this.labH125.BackColor = Color.Green;
						this.labH126.BackColor = Color.Green;
						this.labH124.Text = "00:00";
						this.labH125.Text = "00:00";
						this.labH126.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn42.Visible = true;
					this.labelEn42.BackColor = Color.Red;
					this.labelEn42.ForeColor = Color.White;
					this.labelEn42.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn42.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button42.BackColor = Color.MediumPurple;
					this.labH124.BackColor = Color.MediumPurple;
					this.labH125.BackColor = Color.MediumPurple;
					this.labH126.BackColor = Color.MediumPurple;
					this.labH124.Text = "";
					this.labH125.Text = "";
					this.labH126.Font = new Font(this.labH126.Font.FontFamily, 10);
					this.labH126.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button42.BackColor = Color.DarkSlateBlue;
					this.labH124.BackColor = Color.DarkSlateBlue;
					this.labH125.BackColor = Color.DarkSlateBlue;
					this.labH126.BackColor = Color.DarkSlateBlue;
					this.labH124.Text = "";
					this.labH125.Text = "";
					this.labH126.Font = new Font(this.labH126.Font.FontFamily, 11);
					this.labH126.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button42.BackColor = Color.Turquoise;
					this.labH124.BackColor = Color.Turquoise;
					this.labH125.BackColor = Color.Turquoise;
					this.labH126.BackColor = Color.Turquoise;
					this.button42.ForeColor = Color.Black;
					this.labH124.ForeColor = Color.Black;
					this.labH125.ForeColor = Color.Black;
					this.labH126.ForeColor = Color.Black;
					this.labH124.Text = "";
					this.labH125.Text = "";
					this.labH126.Font = new Font(this.labH126.Font.FontFamily, 10);
					this.labH126.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button42.BackColor = Color.Navy;
					this.labH124.BackColor = Color.Navy;
					this.labH125.BackColor = Color.Navy;
					this.labH126.BackColor = Color.Navy;
					this.button42.ForeColor = Color.White;
					this.labH124.ForeColor = Color.White;
					this.labH125.ForeColor = Color.White;
					this.labH126.ForeColor = Color.White;
					this.labH124.Text = "";
					this.labH125.Text = "";
					this.labH126.Font = new Font(this.labH126.Font.FontFamily, 10);
					this.labH126.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button42.BackColor = Color.OrangeRed;
					this.labH124.BackColor = Color.OrangeRed;
					this.labH125.BackColor = Color.OrangeRed;
					this.labH126.BackColor = Color.OrangeRed;
					this.button42.ForeColor = Color.White;
					this.labH124.ForeColor = Color.White;
					this.labH125.ForeColor = Color.White;
					this.labH126.ForeColor = Color.White;
					this.labH124.Text = "";
					this.labH125.Text = "";
					this.labH126.Font = new Font(this.labH126.Font.FontFamily, 10);
					this.labH126.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button42.BackColor = Color.Orange;
					this.labH124.BackColor = Color.Orange;
					this.labH125.BackColor = Color.Orange;
					this.labH126.BackColor = Color.Orange;
					this.button42.ForeColor = Color.White;
					this.labH124.ForeColor = Color.White;
					this.labH125.ForeColor = Color.White;
					this.labH126.ForeColor = Color.White;
					this.labH124.Text = "";
					this.labH125.Text = "";
					this.labH126.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH126.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h42
			#region h43
			if (NurmdHabitacion==43)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button43.BackColor = Color.SteelBlue;
						this.labH127.BackColor = Color.SteelBlue;
						this.labH128.BackColor = Color.SteelBlue;
						this.labH129.BackColor = Color.SteelBlue;
						this.labH127.Text = Hinicia;
						this.labH128.Text = HFinaliza;
						this.labH129.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else 
					{
						this.button43.BackColor = Color.Red;
						this.labH127.BackColor = Color.Red;
						this.labH128.BackColor = Color.Red;
						this.labH129.BackColor = Color.Red;
						this.labH127.Text = Hinicia;
						this.labH128.Text = HFinaliza;
						this.labH129.Text = "OCUPADO";
					}
				}
				if (e_hab==3)
				{

					{
						this.button43.BackColor = Color.Yellow;
						this.button43.ForeColor = Color.Black;
						this.labH127.BackColor = Color.Yellow;
						this.labH128.BackColor = Color.Yellow;
						this.labH129.BackColor = Color.Yellow;
						this.labH127.ForeColor = Color.Black;
						this.labH128.ForeColor = Color.Black;
						this.labH129.ForeColor = Color.Black;
						this.labH127.Text = Hinicia;
						this.labH128.Text = HFinaliza;
						this.labH129.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button43.BackColor = Color.Silver;
					this.labH127.BackColor = Color.Silver;
					this.labH128.BackColor = Color.Silver;
					this.labH129.BackColor = Color.Silver;
					this.labH129.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button43.BackColor = Color.SaddleBrown;
						this.labH127.BackColor = Color.SaddleBrown;
						this.labH128.BackColor = Color.SaddleBrown;
						this.labH129.BackColor = Color.SaddleBrown;
						this.labH127.Text = hri;
						this.labH128.Text = hrf;
						this.labH129.Text = "RESERVADO";
					}
					else
					{
						this.button43.BackColor = Color.Green;
						this.labH127.BackColor = Color.Green;
						this.labH128.BackColor = Color.Green;
						this.labH129.BackColor = Color.Green;
						this.labH127.Text = "00:00";
						this.labH128.Text = "00:00";
						this.labH129.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn43.Visible = true;
					this.labelEn43.BackColor = Color.Red;
					this.labelEn43.ForeColor = Color.White;
					this.labelEn43.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn43.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button43.BackColor = Color.MediumPurple;
					this.labH127.BackColor = Color.MediumPurple;
					this.labH128.BackColor = Color.MediumPurple;
					this.labH129.BackColor = Color.MediumPurple;
					this.labH127.Text = "";
					this.labH128.Text = "";
					this.labH129.Font = new Font(this.labH129.Font.FontFamily, 10);
					this.labH129.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button43.BackColor = Color.DarkSlateBlue;
					this.labH127.BackColor = Color.DarkSlateBlue;
					this.labH128.BackColor = Color.DarkSlateBlue;
					this.labH129.BackColor = Color.DarkSlateBlue;
					this.labH127.Text = "";
					this.labH128.Text = "";
					this.labH129.Font = new Font(this.labH129.Font.FontFamily, 11);
					this.labH129.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button43.BackColor = Color.Turquoise;
					this.labH127.BackColor = Color.Turquoise;
					this.labH128.BackColor = Color.Turquoise;
					this.labH129.BackColor = Color.Turquoise;
					this.button43.ForeColor = Color.Black;
					this.labH127.ForeColor = Color.Black;
					this.labH128.ForeColor = Color.Black;
					this.labH129.ForeColor = Color.Black;
					this.labH127.Text = "";
					this.labH128.Text = "";
					this.labH129.Font = new Font(this.labH129.Font.FontFamily, 10);
					this.labH129.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button43.BackColor = Color.Navy;
					this.labH127.BackColor = Color.Navy;
					this.labH128.BackColor = Color.Navy;
					this.labH129.BackColor = Color.Navy;
					this.button43.ForeColor = Color.White;
					this.labH127.ForeColor = Color.White;
					this.labH128.ForeColor = Color.White;
					this.labH129.ForeColor = Color.White;
					this.labH127.Text = "";
					this.labH128.Text = "";
					this.labH129.Font = new Font(this.labH129.Font.FontFamily, 10);
					this.labH129.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button43.BackColor = Color.OrangeRed;
					this.labH127.BackColor = Color.OrangeRed;
					this.labH128.BackColor = Color.OrangeRed;
					this.labH129.BackColor = Color.OrangeRed;
					this.button43.ForeColor = Color.White;
					this.labH127.ForeColor = Color.White;
					this.labH128.ForeColor = Color.White;
					this.labH129.ForeColor = Color.White;
					this.labH127.Text = "";
					this.labH128.Text = "";
					this.labH129.Font = new Font(this.labH129.Font.FontFamily, 10);
					this.labH129.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button43.BackColor = Color.Orange;
					this.labH127.BackColor = Color.Orange;
					this.labH128.BackColor = Color.Orange;
					this.labH129.BackColor = Color.Orange;
					this.button43.ForeColor = Color.White;
					this.labH127.ForeColor = Color.White;
					this.labH128.ForeColor = Color.White;
					this.labH129.ForeColor = Color.White;
					this.labH127.Text = "";
					this.labH128.Text = "";
					this.labH129.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH129.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h43
			#region h44
			if (NurmdHabitacion==44)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button44.BackColor = Color.SteelBlue;
						this.labH130.BackColor = Color.SteelBlue;
						this.labH131.BackColor = Color.SteelBlue;
						this.labH132.BackColor = Color.SteelBlue;
						this.labH130.Text = Hinicia;
						this.labH131.Text = HFinaliza;
						this.labH132.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else 
					{
						this.button44.BackColor = Color.Red;
						this.labH130.BackColor = Color.Red;
						this.labH131.BackColor = Color.Red;
						this.labH132.BackColor = Color.Red;
						this.labH130.Text = Hinicia;
						this.labH131.Text = HFinaliza;
						this.labH132.Text = "OCUPADO";
					}
				}
				if (e_hab==3)
				{

					{
						this.button44.BackColor = Color.Yellow;
						this.button44.ForeColor = Color.Black;
						this.labH130.BackColor = Color.Yellow;
						this.labH131.BackColor = Color.Yellow;
						this.labH132.BackColor = Color.Yellow;
						this.labH130.ForeColor = Color.Black;
						this.labH131.ForeColor = Color.Black;
						this.labH132.ForeColor = Color.Black;
						this.labH130.Text = Hinicia;
						this.labH131.Text = HFinaliza;
						this.labH132.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button44.BackColor = Color.Silver;
					this.labH130.BackColor = Color.Silver;
					this.labH131.BackColor = Color.Silver;
					this.labH132.BackColor = Color.Silver;
					this.labH132.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button44.BackColor = Color.SaddleBrown;
						this.labH130.BackColor = Color.SaddleBrown;
						this.labH131.BackColor = Color.SaddleBrown;
						this.labH132.BackColor = Color.SaddleBrown;
						this.labH130.Text = hri;
						this.labH131.Text = hrf;
						this.labH132.Text = "RESERVADO";
					}
					else
					{
						this.button44.BackColor = Color.Green;
						this.labH130.BackColor = Color.Green;
						this.labH131.BackColor = Color.Green;
						this.labH132.BackColor = Color.Green;
						this.labH130.Text = "00:00";
						this.labH131.Text = "00:00";
						this.labH132.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn44.Visible = true;
					this.labelEn44.BackColor = Color.Red;
					this.labelEn44.ForeColor = Color.White;
					this.labelEn44.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn44.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button44.BackColor = Color.MediumPurple;
					this.labH130.BackColor = Color.MediumPurple;
					this.labH131.BackColor = Color.MediumPurple;
					this.labH132.BackColor = Color.MediumPurple;
					this.labH130.Text = "";
					this.labH131.Text = "";
					this.labH132.Font = new Font(this.labH132.Font.FontFamily, 10);
					this.labH132.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button44.BackColor = Color.DarkSlateBlue;
					this.labH130.BackColor = Color.DarkSlateBlue;
					this.labH131.BackColor = Color.DarkSlateBlue;
					this.labH132.BackColor = Color.DarkSlateBlue;
					this.labH130.Text = "";
					this.labH131.Text = "";
					this.labH132.Font = new Font(this.labH132.Font.FontFamily, 11);
					this.labH132.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button44.BackColor = Color.Turquoise;
					this.labH130.BackColor = Color.Turquoise;
					this.labH131.BackColor = Color.Turquoise;
					this.labH132.BackColor = Color.Turquoise;
					this.button44.ForeColor = Color.Black;
					this.labH130.ForeColor = Color.Black;
					this.labH131.ForeColor = Color.Black;
					this.labH132.ForeColor = Color.Black;
					this.labH130.Text = "";
					this.labH131.Text = "";
					this.labH132.Font = new Font(this.labH132.Font.FontFamily, 10);
					this.labH132.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button44.BackColor = Color.Navy;
					this.labH130.BackColor = Color.Navy;
					this.labH131.BackColor = Color.Navy;
					this.labH132.BackColor = Color.Navy;
					this.button44.ForeColor = Color.White;
					this.labH130.ForeColor = Color.White;
					this.labH131.ForeColor = Color.White;
					this.labH132.ForeColor = Color.White;
					this.labH130.Text = "";
					this.labH131.Text = "";
					this.labH132.Font = new Font(this.labH132.Font.FontFamily, 10);
					this.labH132.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button44.BackColor = Color.OrangeRed;
					this.labH130.BackColor = Color.OrangeRed;
					this.labH131.BackColor = Color.OrangeRed;
					this.labH132.BackColor = Color.OrangeRed;
					this.button44.ForeColor = Color.White;
					this.labH130.ForeColor = Color.White;
					this.labH131.ForeColor = Color.White;
					this.labH132.ForeColor = Color.White;
					this.labH130.Text = "";
					this.labH131.Text = "";
					this.labH132.Font = new Font(this.labH132.Font.FontFamily, 10);
					this.labH132.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button44.BackColor = Color.Orange;
					this.labH130.BackColor = Color.Orange;
					this.labH131.BackColor = Color.Orange;
					this.labH132.BackColor = Color.Orange;
					this.button44.ForeColor = Color.White;
					this.labH130.ForeColor = Color.White;
					this.labH131.ForeColor = Color.White;
					this.labH132.ForeColor = Color.White;
					this.labH130.Text = "";
					this.labH131.Text = "";
					this.labH132.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH132.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h44
			#region h45
			if (NurmdHabitacion==45)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button45.BackColor = Color.SteelBlue;
						this.labH133.BackColor = Color.SteelBlue;
						this.labH134.BackColor = Color.SteelBlue;
						this.labH135.BackColor = Color.SteelBlue;
						this.labH133.Text = Hinicia;
						this.labH134.Text = HFinaliza;
						this.labH135.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else 
					{
						this.button45.BackColor = Color.Red;
						this.labH133.BackColor = Color.Red;
						this.labH134.BackColor = Color.Red;
						this.labH135.BackColor = Color.Red;
						this.labH133.Text = Hinicia;
						this.labH134.Text = HFinaliza;
						this.labH135.Text = "OCUPADO";
					}
				}
				if (e_hab==3)
				{

					{
						this.button45.BackColor = Color.Yellow;
						this.button45.ForeColor = Color.Black;
						this.labH133.BackColor = Color.Yellow;
						this.labH134.BackColor = Color.Yellow;
						this.labH135.BackColor = Color.Yellow;
						this.labH133.ForeColor = Color.Black;
						this.labH134.ForeColor = Color.Black;
						this.labH135.ForeColor = Color.Black;
						this.labH133.Text = Hinicia;
						this.labH134.Text = HFinaliza;
						this.labH135.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button45.BackColor = Color.Silver;
					this.labH133.BackColor = Color.Silver;
					this.labH134.BackColor = Color.Silver;
					this.labH135.BackColor = Color.Silver;
					this.labH135.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button45.BackColor = Color.SaddleBrown;
						this.labH133.BackColor = Color.SaddleBrown;
						this.labH134.BackColor = Color.SaddleBrown;
						this.labH135.BackColor = Color.SaddleBrown;
						this.labH133.Text = hri;
						this.labH134.Text = hrf;
						this.labH135.Text = "RESERVADO";
					}
					else
					{
						this.button45.BackColor = Color.Green;
						this.labH133.BackColor = Color.Green;
						this.labH134.BackColor = Color.Green;
						this.labH135.BackColor = Color.Green;
						this.labH133.Text = "00:00";
						this.labH134.Text = "00:00";
						this.labH135.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn45.Visible = true;
					this.labelEn45.BackColor = Color.Red;
					this.labelEn45.ForeColor = Color.White;
					this.labelEn45.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn45.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button1.BackColor = Color.MediumPurple;
					this.labH1.BackColor = Color.MediumPurple;
					this.labH2.BackColor = Color.MediumPurple;
					this.labH3.BackColor = Color.MediumPurple;
					this.labH1.Text = "";
					this.labH2.Text = "";
					this.labH3.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH3.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button1.BackColor = Color.DarkSlateBlue;
					this.labH1.BackColor = Color.DarkSlateBlue;
					this.labH2.BackColor = Color.DarkSlateBlue;
					this.labH3.BackColor = Color.DarkSlateBlue;
					this.labH1.Text = "";
					this.labH2.Text = "";
					this.labH3.Font = new Font(this.labH3.Font.FontFamily, 11);
					this.labH3.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button45.BackColor = Color.Turquoise;
					this.labH133.BackColor = Color.Turquoise;
					this.labH134.BackColor = Color.Turquoise;
					this.labH135.BackColor = Color.Turquoise;
					this.button45.ForeColor = Color.Black;
					this.labH133.ForeColor = Color.Black;
					this.labH134.ForeColor = Color.Black;
					this.labH135.ForeColor = Color.Black;
					this.labH133.Text = "";
					this.labH134.Text = "";
					this.labH135.Font = new Font(this.labH135.Font.FontFamily, 10);
					this.labH135.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button45.BackColor = Color.Navy;
					this.labH133.BackColor = Color.Navy;
					this.labH134.BackColor = Color.Navy;
					this.labH135.BackColor = Color.Navy;
					this.button45.ForeColor = Color.White;
					this.labH133.ForeColor = Color.White;
					this.labH134.ForeColor = Color.White;
					this.labH135.ForeColor = Color.White;
					this.labH133.Text = "";
					this.labH134.Text = "";
					this.labH135.Font = new Font(this.labH135.Font.FontFamily, 10);
					this.labH135.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button45.BackColor = Color.OrangeRed;
					this.labH133.BackColor = Color.OrangeRed;
					this.labH134.BackColor = Color.OrangeRed;
					this.labH135.BackColor = Color.OrangeRed;
					this.button45.ForeColor = Color.White;
					this.labH133.ForeColor = Color.White;
					this.labH134.ForeColor = Color.White;
					this.labH135.ForeColor = Color.White;
					this.labH133.Text = "";
					this.labH134.Text = "";
					this.labH135.Font = new Font(this.labH135.Font.FontFamily, 10);
					this.labH135.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button45.BackColor = Color.Orange;
					this.labH133.BackColor = Color.Orange;
					this.labH134.BackColor = Color.Orange;
					this.labH135.BackColor = Color.Orange;
					this.button45.ForeColor = Color.White;
					this.labH133.ForeColor = Color.White;
					this.labH134.ForeColor = Color.White;
					this.labH135.ForeColor = Color.White;
					this.labH133.Text = "";
					this.labH134.Text = "";
					this.labH135.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH135.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h45
			#region h46
			if (NurmdHabitacion==46)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button46.BackColor = Color.SteelBlue;
						this.labH136.BackColor = Color.SteelBlue;
						this.labH137.BackColor = Color.SteelBlue;
						this.labH138.BackColor = Color.SteelBlue;
						this.labH136.Text = Hinicia;
						this.labH137.Text = HFinaliza;
						this.labH138.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else 
					{
						this.button46.BackColor = Color.Red;
						this.labH136.BackColor = Color.Red;
						this.labH137.BackColor = Color.Red;
						this.labH138.BackColor = Color.Red;
						this.labH136.Text = Hinicia;
						this.labH137.Text = HFinaliza;
						this.labH138.Text = "OCUPADO";
					}
				}
				if (e_hab==3)
				{

					{
						this.button46.BackColor = Color.Yellow;
						this.button46.ForeColor = Color.Black;
						this.labH136.BackColor = Color.Yellow;
						this.labH137.BackColor = Color.Yellow;
						this.labH138.BackColor = Color.Yellow;
						this.labH136.ForeColor = Color.Black;
						this.labH137.ForeColor = Color.Black;
						this.labH138.ForeColor = Color.Black;
						this.labH136.Text = Hinicia;
						this.labH137.Text = HFinaliza;
						this.labH138.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button46.BackColor = Color.Silver;
					this.labH136.BackColor = Color.Silver;
					this.labH137.BackColor = Color.Silver;
					this.labH138.BackColor = Color.Silver;
					this.labH138.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button46.BackColor = Color.SaddleBrown;
						this.labH136.BackColor = Color.SaddleBrown;
						this.labH137.BackColor = Color.SaddleBrown;
						this.labH138.BackColor = Color.SaddleBrown;
						this.labH136.Text = hri;
						this.labH137.Text = hrf;
						this.labH138.Text = "RESERVADO";
					}
					else
					{
						this.button46.BackColor = Color.Green;
						this.labH136.BackColor = Color.Green;
						this.labH137.BackColor = Color.Green;
						this.labH138.BackColor = Color.Green;
						this.labH136.Text = "00:00";
						this.labH137.Text = "00:00";
						this.labH138.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn46.Visible = true;
					this.labelEn46.BackColor = Color.Red;
					this.labelEn46.ForeColor = Color.White;
					this.labelEn46.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn46.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button46.BackColor = Color.MediumPurple;
					this.labH136.BackColor = Color.MediumPurple;
					this.labH137.BackColor = Color.MediumPurple;
					this.labH138.BackColor = Color.MediumPurple;
					this.labH136.Text = "";
					this.labH137.Text = "";
					this.labH138.Font = new Font(this.labH138.Font.FontFamily, 10);
					this.labH138.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button46.BackColor = Color.DarkSlateBlue;
					this.labH136.BackColor = Color.DarkSlateBlue;
					this.labH137.BackColor = Color.DarkSlateBlue;
					this.labH138.BackColor = Color.DarkSlateBlue;
					this.labH136.Text = "";
					this.labH137.Text = "";
					this.labH138.Font = new Font(this.labH138.Font.FontFamily, 11);
					this.labH138.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button46.BackColor = Color.Turquoise;
					this.labH136.BackColor = Color.Turquoise;
					this.labH137.BackColor = Color.Turquoise;
					this.labH138.BackColor = Color.Turquoise;
					this.button46.ForeColor = Color.Black;
					this.labH136.ForeColor = Color.Black;
					this.labH137.ForeColor = Color.Black;
					this.labH138.ForeColor = Color.Black;
					this.labH136.Text = "";
					this.labH137.Text = "";
					this.labH138.Font = new Font(this.labH138.Font.FontFamily, 10);
					this.labH138.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button46.BackColor = Color.Navy;
					this.labH136.BackColor = Color.Navy;
					this.labH137.BackColor = Color.Navy;
					this.labH138.BackColor = Color.Navy;
					this.button46.ForeColor = Color.White;
					this.labH136.ForeColor = Color.White;
					this.labH137.ForeColor = Color.White;
					this.labH138.ForeColor = Color.White;
					this.labH136.Text = "";
					this.labH137.Text = "";
					this.labH138.Font = new Font(this.labH138.Font.FontFamily, 10);
					this.labH138.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button46.BackColor = Color.OrangeRed;
					this.labH136.BackColor = Color.OrangeRed;
					this.labH137.BackColor = Color.OrangeRed;
					this.labH138.BackColor = Color.OrangeRed;
					this.button46.ForeColor = Color.White;
					this.labH136.ForeColor = Color.White;
					this.labH137.ForeColor = Color.White;
					this.labH138.ForeColor = Color.White;
					this.labH136.Text = "";
					this.labH137.Text = "";
					this.labH138.Font = new Font(this.labH138.Font.FontFamily, 10);
					this.labH138.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button46.BackColor = Color.Orange;
					this.labH136.BackColor = Color.Orange;
					this.labH137.BackColor = Color.Orange;
					this.labH138.BackColor = Color.Orange;
					this.button46.ForeColor = Color.White;
					this.labH136.ForeColor = Color.White;
					this.labH137.ForeColor = Color.White;
					this.labH138.ForeColor = Color.White;
					this.labH136.Text = "";
					this.labH137.Text = "";
					this.labH138.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH138.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h46
			#region h47
			if (NurmdHabitacion==47)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button47.BackColor = Color.SteelBlue;
						this.labH139.BackColor = Color.SteelBlue;
						this.labH140.BackColor = Color.SteelBlue;
						this.labH141.BackColor = Color.SteelBlue;
						this.labH139.Text = Hinicia;
						this.labH140.Text = HFinaliza;
						this.labH141.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else 
					{
						this.button47.BackColor = Color.Red;
						this.labH139.BackColor = Color.Red;
						this.labH140.BackColor = Color.Red;
						this.labH141.BackColor = Color.Red;
						this.labH139.Text = Hinicia;
						this.labH140.Text = HFinaliza;
						this.labH141.Text = "OCUPADO";
					}
				}
				if (e_hab==3)
				{

					{
						this.button47.BackColor = Color.Yellow;
						this.button47.ForeColor = Color.Black;
						this.labH139.BackColor = Color.Yellow;
						this.labH140.BackColor = Color.Yellow;
						this.labH141.BackColor = Color.Yellow;
						this.labH139.ForeColor = Color.Black;
						this.labH140.ForeColor = Color.Black;
						this.labH141.ForeColor = Color.Black;
						this.labH139.Text = Hinicia;
						this.labH140.Text = HFinaliza;
						this.labH141.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button47.BackColor = Color.Silver;
					this.labH139.BackColor = Color.Silver;
					this.labH140.BackColor = Color.Silver;
					this.labH141.BackColor = Color.Silver;
					this.labH141.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button47.BackColor = Color.SaddleBrown;
						this.labH139.BackColor = Color.SaddleBrown;
						this.labH140.BackColor = Color.SaddleBrown;
						this.labH141.BackColor = Color.SaddleBrown;
						this.labH139.Text = hri;
						this.labH140.Text = hrf;
						this.labH141.Text = "RESERVADO";
					}
					else
					{
						this.button47.BackColor = Color.Green;
						this.labH139.BackColor = Color.Green;
						this.labH140.BackColor = Color.Green;
						this.labH141.BackColor = Color.Green;
						this.labH139.Text = "00:00";
						this.labH140.Text = "00:00";
						this.labH141.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn47.Visible = true;
					this.labelEn47.BackColor = Color.Red;
					this.labelEn47.ForeColor = Color.White;
					this.labelEn47.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn47.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button47.BackColor = Color.MediumPurple;
					this.labH139.BackColor = Color.MediumPurple;
					this.labH140.BackColor = Color.MediumPurple;
					this.labH141.BackColor = Color.MediumPurple;
					this.labH139.Text = "";
					this.labH140.Text = "";
					this.labH141.Font = new Font(this.labH141.Font.FontFamily, 10);
					this.labH141.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button47.BackColor = Color.DarkSlateBlue;
					this.labH139.BackColor = Color.DarkSlateBlue;
					this.labH140.BackColor = Color.DarkSlateBlue;
					this.labH141.BackColor = Color.DarkSlateBlue;
					this.labH139.Text = "";
					this.labH140.Text = "";
					this.labH141.Font = new Font(this.labH141.Font.FontFamily, 11);
					this.labH141.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button47.BackColor = Color.Turquoise;
					this.labH139.BackColor = Color.Turquoise;
					this.labH140.BackColor = Color.Turquoise;
					this.labH141.BackColor = Color.Turquoise;
					this.button47.ForeColor = Color.Black;
					this.labH139.ForeColor = Color.Black;
					this.labH140.ForeColor = Color.Black;
					this.labH141.ForeColor = Color.Black;
					this.labH139.Text = "";
					this.labH140.Text = "";
					this.labH141.Font = new Font(this.labH141.Font.FontFamily, 10);
					this.labH141.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button47.BackColor = Color.Navy;
					this.labH139.BackColor = Color.Navy;
					this.labH140.BackColor = Color.Navy;
					this.labH141.BackColor = Color.Navy;
					this.button47.ForeColor = Color.White;
					this.labH139.ForeColor = Color.White;
					this.labH140.ForeColor = Color.White;
					this.labH141.ForeColor = Color.White;
					this.labH139.Text = "";
					this.labH140.Text = "";
					this.labH141.Font = new Font(this.labH141.Font.FontFamily, 10);
					this.labH141.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button47.BackColor = Color.OrangeRed;
					this.labH139.BackColor = Color.OrangeRed;
					this.labH140.BackColor = Color.OrangeRed;
					this.labH141.BackColor = Color.OrangeRed;
					this.button47.ForeColor = Color.White;
					this.labH139.ForeColor = Color.White;
					this.labH140.ForeColor = Color.White;
					this.labH141.ForeColor = Color.White;
					this.labH139.Text = "";
					this.labH140.Text = "";
					this.labH141.Font = new Font(this.labH141.Font.FontFamily, 10);
					this.labH141.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button47.BackColor = Color.Orange;
					this.labH139.BackColor = Color.Orange;
					this.labH140.BackColor = Color.Orange;
					this.labH141.BackColor = Color.Orange;
					this.button47.ForeColor = Color.White;
					this.labH139.ForeColor = Color.White;
					this.labH140.ForeColor = Color.White;
					this.labH141.ForeColor = Color.White;
					this.labH139.Text = "";
					this.labH140.Text = "";
					this.labH141.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH141.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h47
			#region h48
			if (NurmdHabitacion==48)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button48.BackColor = Color.SteelBlue;
						this.labH142.BackColor = Color.SteelBlue;
						this.labH143.BackColor = Color.SteelBlue;
						this.labH144.BackColor = Color.SteelBlue;
						this.labH142.Text = Hinicia;
						this.labH143.Text = HFinaliza;
						this.labH144.Text = "PASADO:"+dif_mins.ToString()+" Min.";
					}
					else 
					{
						this.button48.BackColor = Color.Red;
						this.labH142.BackColor = Color.Red;
						this.labH143.BackColor = Color.Red;
						this.labH144.BackColor = Color.Red;
						this.labH142.Text = Hinicia;
						this.labH143.Text = HFinaliza;
						this.labH144.Text = "OCUPADO";
					}
				}
				if (e_hab==3)
				{

					{
						this.button48.BackColor = Color.Yellow;
						this.button48.ForeColor = Color.Black;
						this.labH142.BackColor = Color.Yellow;
						this.labH143.BackColor = Color.Yellow;
						this.labH144.BackColor = Color.Yellow;
						this.labH142.ForeColor = Color.Black;
						this.labH143.ForeColor = Color.Black;
						this.labH144.ForeColor = Color.Black;
						this.labH142.Text = Hinicia;
						this.labH143.Text = HFinaliza;
						this.labH144.Text = "LIMPIEZA";
					}
				}
				if (e_hab==4)
				{
					this.button48.BackColor = Color.Silver;
					this.labH142.BackColor = Color.Silver;
					this.labH143.BackColor = Color.Silver;
					this.labH144.BackColor = Color.Silver;
					this.labH144.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
					{
						this.button48.BackColor = Color.SaddleBrown;
						this.labH142.BackColor = Color.SaddleBrown;
						this.labH143.BackColor = Color.SaddleBrown;
						this.labH144.BackColor = Color.SaddleBrown;
						this.labH142.Text = hri;
						this.labH143.Text = hrf;
						this.labH144.Text = "RESERVADO";
					}
					else
					{
						this.button48.BackColor = Color.Green;
						this.labH142.BackColor = Color.Green;
						this.labH143.BackColor = Color.Green;
						this.labH144.BackColor = Color.Green;
						this.labH142.Text = "00:00";
						this.labH143.Text = "00:00";
						this.labH144.Text = "$ "+dPrecio.ToString()+" LIBRE";
					}
				}
				if (e_enc == 1)
				{
					this.labelEn48.Visible = true;
					this.labelEn48.BackColor = Color.Red;
					this.labelEn48.ForeColor = Color.White;
					this.labelEn48.Text = "ENCUESTA";
				}
				else
				{
					this.labelEn48.Visible = false;
				}
				if (e_hab == 5)
				{
					this.button48.BackColor = Color.MediumPurple;
					this.labH142.BackColor = Color.MediumPurple;
					this.labH143.BackColor = Color.MediumPurple;
					this.labH144.BackColor = Color.MediumPurple;
					this.labH142.Text = "";
					this.labH143.Text = "";
					this.labH144.Font = new Font(this.labH144.Font.FontFamily, 10);
					this.labH144.Text = "INGRESO CLIENTE";
				}

				if (e_hab == 6)
				{
					this.button48.BackColor = Color.DarkSlateBlue;
					this.labH142.BackColor = Color.DarkSlateBlue;
					this.labH143.BackColor = Color.DarkSlateBlue;
					this.labH144.BackColor = Color.DarkSlateBlue;
					this.labH142.Text = "";
					this.labH143.Text = "";
					this.labH144.Font = new Font(this.labH144.Font.FontFamily, 11);
					this.labH144.Text = "ENVÍO COBRO";
				}
				if (e_hab == 7)
				{
					this.button48.BackColor = Color.Turquoise;
					this.labH142.BackColor = Color.Turquoise;
					this.labH143.BackColor = Color.Turquoise;
					this.labH144.BackColor = Color.Turquoise;
					this.button48.ForeColor = Color.Black;
					this.labH142.ForeColor = Color.Black;
					this.labH143.ForeColor = Color.Black;
					this.labH144.ForeColor = Color.Black;
					this.labH142.Text = "";
					this.labH143.Text = "";
					this.labH144.Font = new Font(this.labH144.Font.FontFamily, 10);
					this.labH144.Text = "PROCESO LIMPIEZA";	
				}
				if (e_hab == 8)
				{
					this.button48.BackColor = Color.Navy;
					this.labH142.BackColor = Color.Navy;
					this.labH143.BackColor = Color.Navy;
					this.labH144.BackColor = Color.Navy;
					this.button48.ForeColor = Color.White;
					this.labH142.ForeColor = Color.White;
					this.labH143.ForeColor = Color.White;
					this.labH144.ForeColor = Color.White;
					this.labH142.Text = "";
					this.labH143.Text = "";
					this.labH144.Font = new Font(this.labH144.Font.FontFamily, 10);
					this.labH144.Text = "APROBAR LIMPIEZA";				
				}
				if (e_hab == 9)
				{
					this.button48.BackColor = Color.OrangeRed;
					this.labH142.BackColor = Color.OrangeRed;
					this.labH143.BackColor = Color.OrangeRed;
					this.labH144.BackColor = Color.OrangeRed;
					this.button48.ForeColor = Color.White;
					this.labH142.ForeColor = Color.White;
					this.labH143.ForeColor = Color.White;
					this.labH144.ForeColor = Color.White;
					this.labH142.Text = "";
					this.labH143.Text = "";
					this.labH144.Font = new Font(this.labH144.Font.FontFamily, 10);
					this.labH144.Text = "CORREGIR LIMPIEZA";
				}
				if (e_hab == 10)
				{
					this.button48.BackColor = Color.Orange;
					this.labH142.BackColor = Color.Orange;
					this.labH143.BackColor = Color.Orange;
					this.labH144.BackColor = Color.Orange;
					this.button48.ForeColor = Color.White;
					this.labH142.ForeColor = Color.White;
					this.labH143.ForeColor = Color.White;
					this.labH144.ForeColor = Color.White;
					this.labH142.Text = "";
					this.labH143.Text = "";
					this.labH144.Font = new Font(this.labH3.Font.FontFamily, 10);
					this.labH144.Text = "EN PROCESO DE SALIDA";
				}
			}
			#endregion h48
			#endregion habitacion estado
		}

		static public void ColorearBotones(Form ofrm, int TotHab)		{			int i = 48;			foreach (Control oCtrl in ofrm.Controls)			{                				if (oCtrl is Button)				{					oCtrl.ForeColor = Color.White;					oCtrl.BackColor = Color.Green;					oCtrl.Text = "Habitación # " + i.ToString().Trim();					if (i > TotHab) oCtrl.Visible = false;					i --;				}			}		}

		static public void ColorearEqiquetas(Form ofrm, int TolEtiq)		{			int i = 48*3;			int te = TolEtiq*3;			int y = TolEtiq*3;			foreach (Control oCtrl in ofrm.Controls)			{                				if (oCtrl is Label && !oCtrl.Name.StartsWith("labelEn"))				{//					if(oCtrl.Name == "lblPrecioH1")MessageBox.Show(oCtrl.Text.ToString());					oCtrl.ForeColor = Color.White;          oCtrl.BackColor = Color.Green;					if (i>te) oCtrl.Visible = false;
					else if(y==i){							oCtrl.Text = "LIBRE";							y = i-3;				  }					i --;				}			}		}

		private void AbreFacturaHabitacion(int iNumero, int iBodega)
		{
			int idHabitacion = 0;
			int idCompra = 0;
			int iOcupado = 0;
			bool bServicioPagado = false;

			string sSQL = string.Format("Select idHabitacion, ISNULL(idCompra, 0), Ocupado, ServicioPagado From Habitaciones Where NumeroHabitacion = {0} And Bodega = {1}", iNumero, iBodega);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				idHabitacion = dr.GetInt32(0);
				idCompra = dr.GetInt32(1);
				iOcupado = dr.GetInt32(2);
				bServicioPagado = dr.GetBoolean(3);
			}
			dr.Close();

			if (iOcupado == 3)
			{
				MessageBox.Show("La Habitacion esta en proceso de Limpieza\n\nCambie el estado a Libre para continuar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
				return;
			}

			frmFacturacionHabitaciones FH = new frmFacturacionHabitaciones(idHabitacion, iBodega, true, idCompra, bServicioPagado, iRetTarifa);
			FH.ShowDialog();

			EstadoHabitaciones();
		}

		public void FacturaHabication(int nHab,string nlaBod, int idNumroHab)
		{
			frmFCHabitaciones miFC = new frmFCHabitaciones(nHab,nlaBod,idNumroHab);
			miFC.ShowDialog();

			EstadoHabitaciones();
		}

		public void TarifaExpress(int iNumHabitacion, int iValidaBodega)
		{
			#region Variables
			int iValOcupado = 0;
			int iValEstado = 0;
			#endregion Variables

			string sSQL = string.Format("Select ISNULL(Ocupado, 0) As Ocupado, ISNULL(CONVERT(Int,Estado), 0) As Estado From Habitaciones Where Bodega = {0} And NumeroHabitacion = {1}", iValidaBodega, iNumHabitacion);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				iValOcupado = dr.GetInt32(0);
				iValEstado = dr.GetInt32(1);
			}
			dr.Close();

			#region Tarifa Express
//			if (iValOcupado == 1 && iValEstado == 0)
//			{
//				if (iValidaBodega == 2)
//				{
//					using (frmHabitacionExpres miHE = new frmHabitacionExpres())
//					{
//						if (DialogResult.OK == miHE.ShowDialog())
//						{
//							iRetTarifa = miHE.iTarifa;
//							AbreFacturaHabitacion(iNumHabitacion, iBodega);
//						}
//						else
//						{
//							return;
//						}
//					}
//				}
//				else
//				{
//					iRetTarifa = 0;
//					AbreFacturaHabitacion(iNumHabitacion, iBodega);
//				}
//			}
//			else
//			{
//				iRetTarifa = 0;
//				AbreFacturaHabitacion(iNumHabitacion, iBodega);
//			}
			#endregion Tarifa Express

			iRetTarifa = 0;

			if (MenuLatinium.iNivel == 31)
			{
				if(iValEstado == 0 && iValOcupado == 1)
				{
					AbreFacturaHabitacion(iNumHabitacion, iBodega);
				}
				else
				{
					MessageBox.Show("La Habitacion se encuentra en TEMPORAL", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			else
			{
				if (iValOcupado == 7 || iValOcupado == 8 || iValOcupado == 9)
				{
					return;
				}
				AbreFacturaHabitacion(iNumHabitacion, iBodega);
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(1, iBodega);

//			AbreFacturaHabitacion(1, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=1 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(1,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}
		
		private void button2_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(2, iBodega);

//			AbreFacturaHabitacion(2, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=2 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(2,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(3, iBodega);

//			AbreFacturaHabitacion(3, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=3 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(3,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button4_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(4, iBodega);

//			AbreFacturaHabitacion(4, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=4 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(4,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button5_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(5, iBodega);

//			AbreFacturaHabitacion(5, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=5 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(5,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button6_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(6, iBodega);

//			AbreFacturaHabitacion(6, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=6 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(6,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button7_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(7, iBodega);

//			AbreFacturaHabitacion(7, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=7 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(7,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button8_Click(object sender, System.EventArgs e)
		{
			iRetTarifa = 0;
			TarifaExpress(8, iBodega);

//        AbreFacturaHabitacion(8, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=8 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(8,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button9_Click(object sender, System.EventArgs e)
		{
			iRetTarifa = 0;
			TarifaExpress(9, iBodega);

//			AbreFacturaHabitacion(9, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=9 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(9,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button10_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(10, iBodega);

//			AbreFacturaHabitacion(10, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=10 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(10,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button11_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(11, iBodega);

//			AbreFacturaHabitacion(11, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=11 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(11,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button12_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(12, iBodega);

//			AbreFacturaHabitacion(12, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=12 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(12,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button13_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(13, iBodega);

//			AbreFacturaHabitacion(13, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=13 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(13,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button14_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(14, iBodega);

//			AbreFacturaHabitacion(14, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=14 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(14,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button15_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(15, iBodega);

//			AbreFacturaHabitacion(15, iBodega);
						
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=15 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(15,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button16_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(16, iBodega);

//			AbreFacturaHabitacion(16, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=16 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(16,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button17_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(17, iBodega);

//			AbreFacturaHabitacion(17, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=17 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(17,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button18_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(18, iBodega);

//			AbreFacturaHabitacion(18, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=18 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(18,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button19_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(19, iBodega);

//			AbreFacturaHabitacion(19, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=19 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(19,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button20_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(20, iBodega);

//			AbreFacturaHabitacion(20, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=20 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(20,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button21_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(21, iBodega);

//			AbreFacturaHabitacion(21, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=21 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(21,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button22_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(22, iBodega);

//			AbreFacturaHabitacion(22, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=22 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(22,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button23_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(23, iBodega);

//			AbreFacturaHabitacion(23, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=23 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(23,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button24_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(24, iBodega);

//			AbreFacturaHabitacion(24, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=24 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(24,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button25_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(25, iBodega);

//			AbreFacturaHabitacion(25, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=25 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(25,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button26_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(26, iBodega);

//			AbreFacturaHabitacion(26, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=26 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(26,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button27_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(27, iBodega);

//			AbreFacturaHabitacion(27, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=27 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(27,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button28_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(28, iBodega);

//			AbreFacturaHabitacion(28, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=28 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(28,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button29_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(29, iBodega);

//			AbreFacturaHabitacion(29, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=29 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(29,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button30_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(30, iBodega);

//			AbreFacturaHabitacion(30, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=30 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(30,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button31_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(31, iBodega);

//			AbreFacturaHabitacion(31, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=31 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(31,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button32_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(32, iBodega);

//			AbreFacturaHabitacion(32, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=32 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(32,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button33_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(33, iBodega);

//			AbreFacturaHabitacion(33, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=33 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(33,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button34_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(34, iBodega);

//			AbreFacturaHabitacion(34, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=34 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(34,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button35_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(35, iBodega);

//			AbreFacturaHabitacion(35, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=35 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(35,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button36_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(36, iBodega);

//			AbreFacturaHabitacion(36, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=36 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(36,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button37_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(37, iBodega);

//			AbreFacturaHabitacion(37, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=37 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(37,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button38_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(38, iBodega);

//			AbreFacturaHabitacion(38, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=38 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(38,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button39_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(39, iBodega);

//			AbreFacturaHabitacion(39, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=39 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(39,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button40_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(40, iBodega);

//			AbreFacturaHabitacion(40, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=40 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(40,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button41_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(41, iBodega);

//			AbreFacturaHabitacion(41, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=41 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(41,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button42_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(42, iBodega);

//			AbreFacturaHabitacion(42, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=42 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(42,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button43_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(43, iBodega);

//			AbreFacturaHabitacion(43, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=43 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(43,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button44_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(44, iBodega);

//			AbreFacturaHabitacion(44, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=44 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(44,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button45_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(45, iBodega);

//			AbreFacturaHabitacion(45, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=45 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(45,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button46_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(46, iBodega);

//			AbreFacturaHabitacion(46, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=46 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(46,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button47_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(47, iBodega);

//			AbreFacturaHabitacion(47, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=47 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(47,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button48_Click(object sender, System.EventArgs e)
		{
			TarifaExpress(48, iBodega);
			
//			AbreFacturaHabitacion(48, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=48 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(48,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void timer1_Tick(object sender, System.EventArgs e)
		{
			EstadoHabitaciones();
		}

		int nboton = 0;

		private void MostrarMenuSiCorresponde(Control botonControl, int numeroBoton)
		{
			try
			{
				int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format(
					"Select Ocupado From Habitaciones Where NumeroHabitacion = {0} And Bodega = {1}",
					numeroBoton, iBodega));

				if (iEstado == 2 || iEstado == 7 || iEstado == 8 || iEstado == 9 || iEstado == 10) return;

				menuIngresoCLI.Enabled = !(iEstado == 5 || iEstado == 6);
				menuEnvioCB.Enabled = (iEstado == 5);
				mnuLimpieza.Enabled = true;
				mnuLibre.Enabled = true;

				nboton = numeroBoton;

				this.MnuEstado.Show(botonControl, new Point(3, 30));
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error consultando estado: " + ex.Message);
			}
		}

		private void button1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e) 
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button1, 1);
			}
		}

		private void button2_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button2, 2);
			}
		}

		private void button3_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button3, 3);
			}

		}

		private void button4_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button4, 4);
			}

		}

		private void button5_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button5, 5);
			}

		}

		private void button6_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
			MostrarMenuSiCorresponde(this.button6, 6);
			}

		}

		private void button7_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button7, 7);
			}

		}

		private void button8_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button8, 8);
			}

		}

		private void button9_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button9, 9);
			}
		}

		private void button10_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button10, 10);
			}

		}

		private void button11_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button11, 11);
			}

		}

		private void button12_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button12, 12);
			}

		}

		private void button13_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button13, 13);
			}

		}

		private void button14_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button14, 14);
			}
		}

		private void button15_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button15, 15);
			}
		}

		private void button16_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button16, 16);
			}
		}

		private void button17_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button17, 17);
			}
		}

		private void button18_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button18, 18);
			}
		}

		private void button19_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button19, 19);
			}
		}

		private void button20_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button20, 20);
			}
		}

		private void button21_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button21, 21);
			}
		}

		private void button22_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button22, 22);
			}
		}

		private void button23_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button23, 23);
			}
		}

		private void button24_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button24, 24);
			}
		}

		private void button25_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button25, 25);
			}
		}

		private void button26_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				MostrarMenuSiCorresponde(this.button26, 26);
			}
		}

		private void button27_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button27, 27);
		}

		private void button28_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button28, 28);
		}

		private void button29_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button29, 29);
		}

		private void button30_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button30, 30);
		}

		private void button31_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button31, 31);
		}

		private void button32_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button32, 32);
		}

		private void button33_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button33, 33);
		}

		private void button34_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button34, 34);
		}

		private void button35_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button35, 35);
		}

		private void button36_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button36, 36);
		}

		private void button37_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button37, 37);
		}

		private void button38_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button38, 38);
		}

		private void button39_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button39, 39);
		}

		private void button40_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button40, 40);
		}

		private void button41_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button41, 41);
		}

		private void button42_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button42, 42);
		}

		private void button43_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button43, 43);
		}

		private void button44_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button44, 44);
		}

		private void button45_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button45, 45);
		}

		private void button46_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button46, 46);
		}

		private void button47_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button47, 47);
		}

		private void button48_MouseDown(object sender, MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
				MostrarMenuSiCorresponde(this.button48, 48);
		}

		private void frmCuadriculaHabitaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void mnuLimpieza_Click(object sender, System.EventArgs e)
		{
			try
			{
				int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Ocupado From Habitaciones Where NumeroHabitacion = {0} And Bodega = {1}", nboton, iBodega));

				if (iEstado == 3)
				{
					MessageBox.Show("La habitación ya esta en estado limpieza", "Información", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (iEstado == 2)
				{
					MessageBox.Show("La habitación esta ocupada, para poner en estado limpieza la habitación contabilice o anule la factura", "Información", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				string sSQL = string.Format("Exec Cambia_Estado_Habitacion {0}, {1}, {2}, 0", nboton, 3, idBodega);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				EstadoHabitaciones();
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnuLibre_Click(object sender, System.EventArgs e)
		{			
			try
			{				
				int idCompra = 0;
				int iOcupado = 0;
				bool bServicioPagado = false;

				string sSQL = string.Format("Select ISNULL(idCompra, 0), Ocupado, ServicioPagado From Habitaciones Where NumeroHabitacion = {0} And Bodega = {1}", nboton, iBodega);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					idCompra = dr.GetInt32(0);
					iOcupado = dr.GetInt32(1);
					//bServicioPagado = dr.GetBoolean(2);
				}
				dr.Close();

				if (iOcupado == 1)
				{
					MessageBox.Show("La habitación esta libre", "Información", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (iOcupado == 2 && idCompra > 0)
				{
					MessageBox.Show("La habitación esta ocupada, para liberarla contabilice o anule la factura", "Información", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				string SP_Act = string.Format("Exec Cambia_Estado_Habitacion {0}, {1}, {2}, 0", nboton, 1, idBodega);
				Funcion.EjecutaSQL(cdsSeteoF, SP_Act);

				EstadoHabitaciones();
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}			
		}

		private void timer2_Tick(object sender, System.EventArgs e)
		{
			this.Text = "";

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);
			
			int iTurno = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTurno '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm")));

			this.Text = "Estado de Habitaciones : " + nombrdlBodega + " - Fecha Hora : " + Convert.ToDateTime(this.dtFecha.Value).ToString("dd-MM-yyyy HH:mm:ss") + " - Turno : " + iTurno.ToString();
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void labH1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void labH2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void labH3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label7_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label42_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label34_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label32_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label30_Click(object sender, System.EventArgs e)
		{
		
		}

		int estadoActual = 0 ;
		private void menuIngresoCLI_Click(object sender, System.EventArgs e)
		{
			try
			{
				int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Ocupado From Habitaciones Where NumeroHabitacion = {0} And Bodega = {1}", nboton, iBodega));

				if (iEstado == 5 || iEstado == 6)
				{
					MessageBox.Show("Esta habitación ya se encuentra en proceso.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				string sSQL = string.Format("Exec Cambia_Estado_Habitacion {0}, {1}, {2}, 0", nboton, 5, idBodega);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				estadoActual = 5;

				EstadoHabitaciones();
			}
			catch (Exception ex2)
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void menuEnvioCB_Click(object sender, System.EventArgs e)
		{
			try
			{
				int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Ocupado From Habitaciones Where NumeroHabitacion = {0} And Bodega = {1}", nboton, iBodega));

				if (iEstado != 5)
				{
					MessageBox.Show("Debe seleccionar primero 'Ingreso Cliente'.", "Proceso Incompleto", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				string sSQL = string.Format("Exec Cambia_Estado_Habitacion {0}, {1}, {2}, 0", nboton, 6, idBodega);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				estadoActual = 6;

				EstadoHabitaciones();

				/*

				ParameterFields paramFields = new ParameterFields();

				ParameterField paramFieldBodega = new ParameterField();
				ParameterDiscreteValue discreteValBodega = new ParameterDiscreteValue();
				paramFieldBodega.ParameterFieldName = "@IdBodega";
				discreteValBodega.Value = iBodega;
				paramFieldBodega.CurrentValues.Add(discreteValBodega);
				paramFields.Add(paramFieldBodega);

				ParameterField paramFieldHabitacion = new ParameterField();
				ParameterDiscreteValue discreteValHabitacion = new ParameterDiscreteValue();
				paramFieldHabitacion.ParameterFieldName = "@Habitacion";
				discreteValHabitacion.Value = nboton;
				paramFieldHabitacion.CurrentValues.Add(discreteValHabitacion);
				paramFields.Add(paramFieldHabitacion);

				Reporte reporte = new Reporte("ReciboCobro.rpt", "");
				reporte.MdiParent = this.MdiParent;
				reporte.Titulo("Recibo de Cobro");
				reporte.crVista.ParameterFieldInfo = paramFields;
				reporte.Show();
				*/
				#region Nombre Impresora
								
				string sImpresora = "";
				sImpresora = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" exec ConsultarBodegaProservicios {0}, 1 ",iBodega));		
			
				//sImpresora = "SLK-TS100";
				#endregion Nombre Impresora

				#region Impresion Formato

				string sReporte = "ReciboCobro.rpt";
				ReportDocument Report1 = new ReportDocument();
				Report1.Load(MenuLatinium.stDirReportes + sReporte);

				ConnectionInfo crConnectionInfo = new ConnectionInfo();
				crConnectionInfo.ServerName = MenuLatinium.stDirServidor;
				crConnectionInfo.DatabaseName = MenuLatinium.stDirNumero;
				crConnectionInfo.UserID = "infoelect";
				crConnectionInfo.Password = "Bl45o6$9";

				Tables crTables = Report1.Database.Tables;
				for (int i = 0; i < crTables.Count; i++)
				{
					Table crTable = crTables[i];
					try
					{
						TableLogOnInfo crTableLogOnInfo = crTable.LogOnInfo;
						crTableLogOnInfo.ConnectionInfo = crConnectionInfo;
						crTable.ApplyLogOnInfo(crTableLogOnInfo);
					}
					catch (Exception exTabla)
					{
						throw new Exception(string.Format("Error aplicando conexión a tabla {0}: {1}", crTable.Name, exTabla.Message), exTabla);
					}
				}

				if (sImpresora != null && sImpresora != "")
				{
					Report1.PrintOptions.PrinterName = sImpresora;
				}
				else
				{
					throw new Exception("Nombre de impresora vacío o inválido.");
				}

				Report1.SetParameterValue("@IdBodega", iBodega);
				Report1.SetParameterValue("@Habitacion", nboton);

				Report1.PrintToPrinter(2, false, 0, 0);

				#endregion Impresion Formato
			

		
				//TarifaExpress(nboton, iBodega);
			}
			catch (Exception ex2)
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
