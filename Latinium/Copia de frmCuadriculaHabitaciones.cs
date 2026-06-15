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
		private System.Windows.Forms.Timer timer1;
		private System.Windows.Forms.ContextMenu MnuEstado;
		int idBodega = 0;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.MenuItem mnuLimpieza;
		private System.Windows.Forms.MenuItem mnuLibre;
		private System.Windows.Forms.Timer timer2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;		
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
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.timer2 = new System.Windows.Forms.Timer(this.components);
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
																																							this.mnuLibre});
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
			// frmCuadriculaHabitaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1232, 542);
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
			timer1.Interval = 120000; //Dos minutos
			timer1.Start();
			timer1.Enabled = true;

			int sw = 0;

			string sSQL = string.Format("Select Count(*) From UsuariosBodegas Where idUsuario = {0}", MenuLatinium.IdUsuario);
			int CantEmpUser = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			if(CantEmpUser>1) 
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
			int difmints = 0;
			string hrini = "";
			string hrfin = "";

			//string lQry = "SELECT a.NumeroHabitacion,IsNull(a.HoraInicia,GetDate()) as HoraInicia,CONVERT(varchar(5),a.HoraInicia,8) as HIn,";
			//lQry += "HoraFinaliza,CONVERT(varchar(5),a.HoraFinaliza,8) as HFi,a.Ocupado,DateDiff(Minute,a.HoraFinaliza,GetDate()) as DifeMins,";
			//lQry += "DateDiff(HOUR,GETDATE(),b.FechaInicia) as TmpoReserva ";
			//lQry += "From Habitaciones a ";
			//lQry += "Left Join ReservaHabitacion b on a.idHabitacion=b.idHabitacion";
			//lQry += "Where Convert(Date,b.FechaInicia)=CONVERT(Date,GETDATE()) and a.idSucursal="+idBodega.ToString();
			//SqlDataReader drEdoHab = Funcion.rEscalarSQL(cdsSeteoF, lQry, true);

			Funcion.EjecutaSQL(cdsSeteoF, 
				string.Format("Update Habitaciones Set Ocupado = 1, Estado = 0, NumeroDocumento = '', idCompra = 0 Where Ocupado = 3 And DATEDIFF(MINUTE, HoraInicia, GETDATE()) > 5 And Bodega = {0}", 
				iBodega));

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

				CambiaColorHabitacion(nmHabtcn,hini,hfin,edoHabit,difmints,TiempoReser,hrini,hrfin);
			}
			drEdoHab.Close();
		}

		public void CambiaColorHabitacion(int NurmdHabitacion, string Hinicia, string HFinaliza, int e_hab, int dif_mins, int TR,string hri, string hrf)
		{
			#region habitaciones estado
			#region h1
				if (NurmdHabitacion==1)
				{
					if (e_hab==2)
					{
						if (dif_mins>0)
						{
							this.button1.BackColor = Color.SteelBlue;
							this.labH1.BackColor = Color.SteelBlue;
							this.labH2.BackColor = Color.SteelBlue;
							this.labH3.BackColor = Color.SteelBlue;
							this.labH1.Text = Hinicia;
							this.labH2.Text = HFinaliza;
							this.labH3.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
					if (e_hab==3)
					{
						if (dif_mins>=5)
						{
							this.button1.BackColor = Color.Green;
							this.labH1.BackColor = Color.Green;
							this.labH2.BackColor = Color.Green;
							this.labH3.BackColor = Color.Green;
							this.labH1.Text = "00:00";
							this.labH2.Text = "00:00";
							this.labH3.Text = "LIBRE";
							/*Actualiza habitacion a estado libre*/
//							string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},1",idBodega);
//							Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
						}
						else
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
							this.labH3.Text = "LIBRE";
						}
					}
				}
			#endregion h1
			#region h2
				if (NurmdHabitacion==2)
				{
					if (e_hab==2)
					{
						if (dif_mins>0)
						{
							this.button2.BackColor = Color.SteelBlue;
							this.labH4.BackColor = Color.SteelBlue;
							this.labH5.BackColor = Color.SteelBlue;
							this.labH6.BackColor = Color.SteelBlue;
							this.labH4.Text = Hinicia;
							this.labH5.Text = HFinaliza;
							this.labH6.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
					if (e_hab==3)
					{
						if (dif_mins>=5)
						{
							this.button2.BackColor = Color.Green;
							this.labH4.BackColor = Color.Green;
							this.labH5.BackColor = Color.Green;
							this.labH6.BackColor = Color.Green;
							this.labH4.Text = "00:00";
							this.labH5.Text = "00:00";
							this.labH6.Text = "LIBRE";
							/*Actualiza habitacion a estado libre*/
//							string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},2",idBodega);
//							//string Qry = string.Format("Exec CambiaEdo_Mantenimiento "+idBodega.ToString()+",2");
//							Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
						}
						else
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
							this.labH6.Text = "LIBRE";
						}
					}
				}
			#endregion h2
			#region h3
			if (NurmdHabitacion==3)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button3.BackColor = Color.SteelBlue;
						this.labH7.BackColor = Color.SteelBlue;
						this.labH8.BackColor = Color.SteelBlue;
						this.labH9.BackColor = Color.SteelBlue;
						this.labH7.Text = Hinicia;
						this.labH8.Text = HFinaliza;
						this.labH9.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button3.BackColor = Color.Green;
						this.labH7.BackColor = Color.Green;
						this.labH8.BackColor = Color.Green;
						this.labH9.BackColor = Color.Green;
						this.labH7.Text = "00:00";
						this.labH8.Text = "00:00";
						this.labH9.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},3",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
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
				if (e_hab==4)
				{
					this.button3.BackColor = Color.Silver;
					this.labH7.BackColor = Color.Silver;
					this.labH8.BackColor = Color.Silver;
					this.labH9.BackColor = Color.Silver;
					this.labH9.Text = "MANTENIMIENTO";
				}
				if (e_hab==1)
				{
					if (TR>0)
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
						this.labH9.Text = "LIBRE";
					}
				}
			}
			#endregion h3
			#region h4
			if (NurmdHabitacion==4)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button4.BackColor = Color.SteelBlue;
						this.labH10.BackColor = Color.SteelBlue;
						this.labH11.BackColor = Color.SteelBlue;
						this.labH12.BackColor = Color.SteelBlue;
						this.labH10.Text = Hinicia;
						this.labH11.Text = HFinaliza;
						this.labH12.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button4.BackColor = Color.Green;
						this.labH10.BackColor = Color.Green;
						this.labH11.BackColor = Color.Green;
						this.labH12.BackColor = Color.Green;
						this.labH10.Text = "00:00";
						this.labH11.Text = "00:00";
						this.labH12.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},4",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH12.Text = "LIBRE";
					}
				}
			}
			#endregion h4
			#region h5
			if (NurmdHabitacion==5)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button5.BackColor = Color.SteelBlue;
						this.labH13.BackColor = Color.SteelBlue;
						this.labH14.BackColor = Color.SteelBlue;
						this.labH15.BackColor = Color.SteelBlue;
						this.labH13.Text = Hinicia;
						this.labH14.Text = HFinaliza;
						this.labH15.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button5.BackColor = Color.Green;
						this.labH13.BackColor = Color.Green;
						this.labH14.BackColor = Color.Green;
						this.labH15.BackColor = Color.Green;
						this.labH13.Text = "00:00";
						this.labH14.Text = "00:00";
						this.labH15.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},5",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH15.Text = "LIBRE";
					}
				}
			}
			#endregion h5
			#region h6
			if (NurmdHabitacion==6)
				{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button6.BackColor = Color.SteelBlue;
						this.labH16.BackColor = Color.SteelBlue;
						this.labH17.BackColor = Color.SteelBlue;
						this.labH18.BackColor = Color.SteelBlue;
						this.labH16.Text = Hinicia;
						this.labH17.Text = HFinaliza;
						this.labH18.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button6.BackColor = Color.Green;
						this.labH16.BackColor = Color.Green;
						this.labH17.BackColor = Color.Green;
						this.labH18.BackColor = Color.Green;
						this.labH16.Text = "00:00";
						this.labH17.Text = "00:00";
						this.labH18.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},6",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH18.Text = "LIBRE";
					}
				}
			}
			#endregion h6
			#region h7
			if (NurmdHabitacion==7)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button7.BackColor = Color.SteelBlue;
						this.labH19.BackColor = Color.SteelBlue;
						this.labH20.BackColor = Color.SteelBlue;
						this.labH21.BackColor = Color.SteelBlue;
						this.labH19.Text = Hinicia;
						this.labH20.Text = HFinaliza;
						this.labH21.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button7.BackColor = Color.Green;
						this.labH19.BackColor = Color.Green;
						this.labH20.BackColor = Color.Green;
						this.labH21.BackColor = Color.Green;
						this.labH19.Text = "00:00";
						this.labH20.Text = "00:00";
						this.labH21.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},7",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH21.Text = "LIBRE";
					}
				}
			}
			#endregion h7
			#region h8
			if (NurmdHabitacion==8)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button8.BackColor = Color.SteelBlue;
						this.labH22.BackColor = Color.SteelBlue;
						this.labH23.BackColor = Color.SteelBlue;
						this.labH24.BackColor = Color.SteelBlue;
						this.labH22.Text = Hinicia;
						this.labH23.Text = HFinaliza;
						this.labH24.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button8.BackColor = Color.Green;
						this.labH22.BackColor = Color.Green;
						this.labH23.BackColor = Color.Green;
						this.labH24.BackColor = Color.Green;
						this.labH22.Text = "00:00";
						this.labH23.Text = "00:00";
						this.labH24.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},8",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH24.Text = "LIBRE";
					}
				}
			}
			#endregion h8
			#region h9
			if (NurmdHabitacion==9)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button9.BackColor = Color.SteelBlue;
						this.labH25.BackColor = Color.SteelBlue;
						this.labH26.BackColor = Color.SteelBlue;
						this.labH27.BackColor = Color.SteelBlue;
						this.labH25.Text = Hinicia;
						this.labH26.Text = HFinaliza;
						this.labH27.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button9.BackColor = Color.Green;
						this.labH25.BackColor = Color.Green;
						this.labH26.BackColor = Color.Green;
						this.labH27.BackColor = Color.Green;
						this.labH25.Text = "00:00";
						this.labH26.Text = "00:00";
						this.labH27.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},9",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH27.Text = "LIBRE";
					}
				}
			}
			#endregion h9
			#region h10
			if (NurmdHabitacion==10)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button10.BackColor = Color.SteelBlue;
						this.labH28.BackColor = Color.SteelBlue;
						this.labH29.BackColor = Color.SteelBlue;
						this.labH30.BackColor = Color.SteelBlue;
						this.labH28.Text = Hinicia;
						this.labH29.Text = HFinaliza;
						this.labH30.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button10.BackColor = Color.Green;
						this.labH28.BackColor = Color.Green;
						this.labH29.BackColor = Color.Green;
						this.labH30.BackColor = Color.Green;
						this.labH28.Text = "00:00";
						this.labH29.Text = "00:00";
						this.labH30.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},10",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH30.Text = "LIBRE";
					}
				}
			}
			#endregion h10
			#region h11
			if (NurmdHabitacion==11)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button11.BackColor = Color.SteelBlue;
						this.labH31.BackColor = Color.SteelBlue;
						this.labH32.BackColor = Color.SteelBlue;
						this.labH33.BackColor = Color.SteelBlue;
						this.labH31.Text = Hinicia;
						this.labH32.Text = HFinaliza;
						this.labH33.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button11.BackColor = Color.Green;
						this.labH31.BackColor = Color.Green;
						this.labH32.BackColor = Color.Green;
						this.labH33.BackColor = Color.Green;
						this.labH31.Text = "00:00";
						this.labH32.Text = "00:00";
						this.labH33.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},11",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH33.Text = "LIBRE";
					}
				}
			}
			#endregion h11
			#region h12
			if (NurmdHabitacion==12)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button12.BackColor = Color.SteelBlue;
						this.labH34.BackColor = Color.SteelBlue;
						this.labH35.BackColor = Color.SteelBlue;
						this.labH36.BackColor = Color.SteelBlue;
						this.labH34.Text = Hinicia;
						this.labH35.Text = HFinaliza;
						this.labH36.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button12.BackColor = Color.Green;
						this.labH34.BackColor = Color.Green;
						this.labH35.BackColor = Color.Green;
						this.labH36.BackColor = Color.Green;
						this.labH34.Text = "00:00";
						this.labH35.Text = "00:00";
						this.labH36.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},12",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH36.Text = "LIBRE";
					}
				}
			}
			#endregion h12
			#region h13
			if (NurmdHabitacion==13)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button13.BackColor = Color.SteelBlue;
						this.labH37.BackColor = Color.SteelBlue;
						this.labH38.BackColor = Color.SteelBlue;
						this.labH39.BackColor = Color.SteelBlue;
						this.labH37.Text = Hinicia;
						this.labH38.Text = HFinaliza;
						this.labH39.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button13.BackColor = Color.Green;
						this.labH37.BackColor = Color.Green;
						this.labH38.BackColor = Color.Green;
						this.labH39.BackColor = Color.Green;
						this.labH37.Text = "00:00";
						this.labH38.Text = "00:00";
						this.labH39.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},13",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH39.Text = "LIBRE";
					}
				}
			}
			#endregion h13
			#region h14
			if (NurmdHabitacion==14)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button14.BackColor = Color.SteelBlue;
						this.labH40.BackColor = Color.SteelBlue;
						this.labH41.BackColor = Color.SteelBlue;
						this.labH42.BackColor = Color.SteelBlue;
						this.labH40.Text = Hinicia;
						this.labH41.Text = HFinaliza;
						this.labH42.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button14.BackColor = Color.Green;
						this.labH40.BackColor = Color.Green;
						this.labH41.BackColor = Color.Green;
						this.labH42.BackColor = Color.Green;
						this.labH40.Text = "00:00";
						this.labH41.Text = "00:00";
						this.labH42.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},14",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH42.Text = "LIBRE";
					}
				}
			}
			#endregion h14
			#region h15
			if (NurmdHabitacion==15)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button15.BackColor = Color.SteelBlue;
						this.labH43.BackColor = Color.SteelBlue;
						this.labH44.BackColor = Color.SteelBlue;
						this.labH45.BackColor = Color.SteelBlue;
						this.labH43.Text = Hinicia;
						this.labH44.Text = HFinaliza;
						this.labH45.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button15.BackColor = Color.Green;
						this.labH43.BackColor = Color.Green;
						this.labH44.BackColor = Color.Green;
						this.labH45.BackColor = Color.Green;
						this.labH43.Text = "00:00";
						this.labH44.Text = "00:00";
						this.labH45.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},15",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH45.Text = "LIBRE";
					}
				}
			}
			#endregion h15
			#region h16
			if (NurmdHabitacion==16)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button16.BackColor = Color.SteelBlue;
						this.labH46.BackColor = Color.SteelBlue;
						this.labH47.BackColor = Color.SteelBlue;
						this.labH48.BackColor = Color.SteelBlue;
						this.labH46.Text = Hinicia;
						this.labH47.Text = HFinaliza;
						this.labH48.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button16.BackColor = Color.Green;
						this.labH46.BackColor = Color.Green;
						this.labH47.BackColor = Color.Green;
						this.labH48.BackColor = Color.Green;
						this.labH46.Text = "00:00";
						this.labH47.Text = "00:00";
						this.labH48.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},16",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH48.Text = "LIBRE";
					}
				}
			}
			#endregion h16
			#region h17
			if (NurmdHabitacion==17)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button17.BackColor = Color.SteelBlue;
						this.labH49.BackColor = Color.SteelBlue;
						this.labH50.BackColor = Color.SteelBlue;
						this.labH51.BackColor = Color.SteelBlue;
						this.labH49.Text = Hinicia;
						this.labH50.Text = HFinaliza;
						this.labH51.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button17.BackColor = Color.Green;
						this.labH49.BackColor = Color.Green;
						this.labH50.BackColor = Color.Green;
						this.labH51.BackColor = Color.Green;
						this.labH49.Text = "00:00";
						this.labH50.Text = "00:00";
						this.labH51.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},17",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH51.Text = "LIBRE";
					}
				}
			}
			#endregion h17
			#region h18
			if (NurmdHabitacion==18)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button18.BackColor = Color.SteelBlue;
						this.labH52.BackColor = Color.SteelBlue;
						this.labH53.BackColor = Color.SteelBlue;
						this.labH54.BackColor = Color.SteelBlue;
						this.labH52.Text = Hinicia;
						this.labH53.Text = HFinaliza;
						this.labH54.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button18.BackColor = Color.Green;
						this.labH52.BackColor = Color.Green;
						this.labH53.BackColor = Color.Green;
						this.labH54.BackColor = Color.Green;
						this.labH52.Text = "00:00";
						this.labH53.Text = "00:00";
						this.labH54.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},18",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH54.Text = "LIBRE";
					}
				}
			}
			#endregion h18
			#region h19
			if (NurmdHabitacion==19)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button19.BackColor = Color.SteelBlue;
						this.labH55.BackColor = Color.SteelBlue;
						this.labH56.BackColor = Color.SteelBlue;
						this.labH57.BackColor = Color.SteelBlue;
						this.labH55.Text = Hinicia;
						this.labH56.Text = HFinaliza;
						this.labH57.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button19.BackColor = Color.Green;
						this.labH55.BackColor = Color.Green;
						this.labH56.BackColor = Color.Green;
						this.labH57.BackColor = Color.Green;
						this.labH55.Text = "00:00";
						this.labH56.Text = "00:00";
						this.labH57.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},19",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH57.Text = "LIBRE";
					}
				}
			}
			#endregion h19
			#region h20
			if (NurmdHabitacion==20)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button20.BackColor = Color.SteelBlue;
						this.labH58.BackColor = Color.SteelBlue;
						this.labH59.BackColor = Color.SteelBlue;
						this.labH60.BackColor = Color.SteelBlue;
						this.labH58.Text = Hinicia;
						this.labH59.Text = HFinaliza;
						this.labH60.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button20.BackColor = Color.Green;
						this.labH58.BackColor = Color.Green;
						this.labH59.BackColor = Color.Green;
						this.labH60.BackColor = Color.Green;
						this.labH58.Text = "00:00";
						this.labH59.Text = "00:00";
						this.labH60.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},20",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH60.Text = "LIBRE";
					}
				}
			}
			#endregion h20
			#region h21
			if (NurmdHabitacion==21)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button21.BackColor = Color.SteelBlue;
						this.labH61.BackColor = Color.SteelBlue;
						this.labH62.BackColor = Color.SteelBlue;
						this.labH63.BackColor = Color.SteelBlue;
						this.labH61.Text = Hinicia;
						this.labH62.Text = HFinaliza;
						this.labH63.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button21.BackColor = Color.Green;
						this.labH61.BackColor = Color.Green;
						this.labH62.BackColor = Color.Green;
						this.labH63.BackColor = Color.Green;
						this.labH61.Text = "00:00";
						this.labH62.Text = "00:00";
						this.labH63.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},21",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH63.Text = "LIBRE";
					}
				}
			}
			#endregion h21
			#region h22
			if (NurmdHabitacion==22)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button22.BackColor = Color.SteelBlue;
						this.labH64.BackColor = Color.SteelBlue;
						this.labH65.BackColor = Color.SteelBlue;
						this.labH66.BackColor = Color.SteelBlue;
						this.labH64.Text = Hinicia;
						this.labH65.Text = HFinaliza;
						this.labH66.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button22.BackColor = Color.Green;
						this.labH64.BackColor = Color.Green;
						this.labH65.BackColor = Color.Green;
						this.labH66.BackColor = Color.Green;
						this.labH64.Text = "00:00";
						this.labH65.Text = "00:00";
						this.labH66.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},22",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH66.Text = "LIBRE";
					}
				}
			}
			#endregion h22
			#region h23
			if (NurmdHabitacion==23)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button23.BackColor = Color.SteelBlue;
						this.labH67.BackColor = Color.SteelBlue;
						this.labH68.BackColor = Color.SteelBlue;
						this.labH69.BackColor = Color.SteelBlue;
						this.labH67.Text = Hinicia;
						this.labH68.Text = HFinaliza;
						this.labH69.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button23.BackColor = Color.Green;
						this.labH67.BackColor = Color.Green;
						this.labH68.BackColor = Color.Green;
						this.labH69.BackColor = Color.Green;
						this.labH67.Text = "00:00";
						this.labH68.Text = "00:00";
						this.labH69.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},23",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH69.Text = "LIBRE";
					}
				}
			}
			#endregion h23
			#region h24
			if (NurmdHabitacion==24)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button24.BackColor = Color.SteelBlue;
						this.labH70.BackColor = Color.SteelBlue;
						this.labH71.BackColor = Color.SteelBlue;
						this.labH72.BackColor = Color.SteelBlue;
						this.labH70.Text = Hinicia;
						this.labH71.Text = HFinaliza;
						this.labH72.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button24.BackColor = Color.Green;
						this.labH70.BackColor = Color.Green;
						this.labH71.BackColor = Color.Green;
						this.labH72.BackColor = Color.Green;
						this.labH70.Text = "00:00";
						this.labH71.Text = "00:00";
						this.labH72.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},24",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH72.Text = "LIBRE";
					}
				}
			}
			#endregion h24
			#region h25
			if (NurmdHabitacion==25)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button25.BackColor = Color.SteelBlue;
						this.labH73.BackColor = Color.SteelBlue;
						this.labH74.BackColor = Color.SteelBlue;
						this.labH75.BackColor = Color.SteelBlue;
						this.labH73.Text = Hinicia;
						this.labH74.Text = HFinaliza;
						this.labH75.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button25.BackColor = Color.Green;
						this.labH73.BackColor = Color.Green;
						this.labH74.BackColor = Color.Green;
						this.labH75.BackColor = Color.Green;
						this.labH73.Text = "00:00";
						this.labH74.Text = "00:00";
						this.labH75.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},25",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH75.Text = "LIBRE";
					}
				}
			}
			#endregion h25
			#region h26
			if (NurmdHabitacion==26)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button26.BackColor = Color.SteelBlue;
						this.labH76.BackColor = Color.SteelBlue;
						this.labH77.BackColor = Color.SteelBlue;
						this.labH78.BackColor = Color.SteelBlue;
						this.labH76.Text = Hinicia;
						this.labH77.Text = HFinaliza;
						this.labH78.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button26.BackColor = Color.Green;
						this.labH76.BackColor = Color.Green;
						this.labH77.BackColor = Color.Green;
						this.labH78.BackColor = Color.Green;
						this.labH76.Text = "00:00";
						this.labH77.Text = "00:00";
						this.labH78.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},26",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH78.Text = "LIBRE";
					}
				}
			}
			#endregion h26
			#region h27
			if (NurmdHabitacion==27)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button27.BackColor = Color.SteelBlue;
						this.labH79.BackColor = Color.SteelBlue;
						this.labH80.BackColor = Color.SteelBlue;
						this.labH81.BackColor = Color.SteelBlue;
						this.labH79.Text = Hinicia;
						this.labH80.Text = HFinaliza;
						this.labH81.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button27.BackColor = Color.Green;
						this.labH79.BackColor = Color.Green;
						this.labH80.BackColor = Color.Green;
						this.labH81.BackColor = Color.Green;
						this.labH79.Text = "00:00";
						this.labH80.Text = "00:00";
						this.labH81.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},27",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH81.Text = "LIBRE";
					}
				}
			}
			#endregion h27
			#region h28
			if (NurmdHabitacion==28)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button28.BackColor = Color.SteelBlue;
						this.labH82.BackColor = Color.SteelBlue;
						this.labH83.BackColor = Color.SteelBlue;
						this.labH84.BackColor = Color.SteelBlue;
						this.labH82.Text = Hinicia;
						this.labH83.Text = HFinaliza;
						this.labH84.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button28.BackColor = Color.Green;
						this.labH82.BackColor = Color.Green;
						this.labH83.BackColor = Color.Green;
						this.labH84.BackColor = Color.Green;
						this.labH82.Text = "00:00";
						this.labH83.Text = "00:00";
						this.labH84.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},28",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH84.Text = "LIBRE";
					}
				}
			}
			#endregion h28
			#region h29
			if (NurmdHabitacion==29)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button29.BackColor = Color.SteelBlue;
						this.labH85.BackColor = Color.SteelBlue;
						this.labH86.BackColor = Color.SteelBlue;
						this.labH87.BackColor = Color.SteelBlue;
						this.labH85.Text = Hinicia;
						this.labH86.Text = HFinaliza;
						this.labH87.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button29.BackColor = Color.Green;
						this.labH85.BackColor = Color.Green;
						this.labH86.BackColor = Color.Green;
						this.labH87.BackColor = Color.Green;
						this.labH85.Text = "00:00";
						this.labH86.Text = "00:00";
						this.labH87.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},29",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH87.Text = "LIBRE";
					}
				}
			}
			#endregion h29
			#region h30
			if (NurmdHabitacion==30)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button30.BackColor = Color.SteelBlue;
						this.labH88.BackColor = Color.SteelBlue;
						this.labH89.BackColor = Color.SteelBlue;
						this.labH90.BackColor = Color.SteelBlue;
						this.labH88.Text = Hinicia;
						this.labH89.Text = HFinaliza;
						this.labH90.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button30.BackColor = Color.Green;
						this.labH88.BackColor = Color.Green;
						this.labH89.BackColor = Color.Green;
						this.labH90.BackColor = Color.Green;
						this.labH88.Text = "00:00";
						this.labH89.Text = "00:00";
						this.labH90.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},30",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH90.Text = "LIBRE";
					}
				}
			}
			#endregion h30
			#region h31
			if (NurmdHabitacion==31)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button31.BackColor = Color.SteelBlue;
						this.labH91.BackColor = Color.SteelBlue;
						this.labH92.BackColor = Color.SteelBlue;
						this.labH93.BackColor = Color.SteelBlue;
						this.labH91.Text = Hinicia;
						this.labH92.Text = HFinaliza;
						this.labH93.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button31.BackColor = Color.Green;
						this.labH91.BackColor = Color.Green;
						this.labH92.BackColor = Color.Green;
						this.labH93.BackColor = Color.Green;
						this.labH91.Text = "00:00";
						this.labH92.Text = "00:00";
						this.labH93.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},31",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH93.Text = "LIBRE";
					}
				}
			}
			#endregion h31
			#region h32
			if (NurmdHabitacion==32)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button32.BackColor = Color.SteelBlue;
						this.labH94.BackColor = Color.SteelBlue;
						this.labH95.BackColor = Color.SteelBlue;
						this.labH96.BackColor = Color.SteelBlue;
						this.labH94.Text = Hinicia;
						this.labH95.Text = HFinaliza;
						this.labH96.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button32.BackColor = Color.Green;
						this.labH94.BackColor = Color.Green;
						this.labH95.BackColor = Color.Green;
						this.labH96.BackColor = Color.Green;
						this.labH94.Text = "00:00";
						this.labH95.Text = "00:00";
						this.labH96.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},32",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH96.Text = "LIBRE";
					}
				}
			}
			#endregion h32
			#region h33
			if (NurmdHabitacion==33)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button33.BackColor = Color.SteelBlue;
						this.labH97.BackColor = Color.SteelBlue;
						this.labH98.BackColor = Color.SteelBlue;
						this.labH99.BackColor = Color.SteelBlue;
						this.labH97.Text = Hinicia;
						this.labH98.Text = HFinaliza;
						this.labH99.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button33.BackColor = Color.Green;
						this.labH97.BackColor = Color.Green;
						this.labH98.BackColor = Color.Green;
						this.labH99.BackColor = Color.Green;
						this.labH97.Text = "00:00";
						this.labH98.Text = "00:00";
						this.labH99.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},33",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH99.Text = "LIBRE";
					}
				}
			}
			#endregion h33
			#region h34
			if (NurmdHabitacion==34)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button34.BackColor = Color.SteelBlue;
						this.labH100.BackColor = Color.SteelBlue;
						this.labH101.BackColor = Color.SteelBlue;
						this.labH102.BackColor = Color.SteelBlue;
						this.labH100.Text = Hinicia;
						this.labH101.Text = HFinaliza;
						this.labH102.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button34.BackColor = Color.Green;
						this.labH100.BackColor = Color.Green;
						this.labH101.BackColor = Color.Green;
						this.labH102.BackColor = Color.Green;
						this.labH100.Text = "00:00";
						this.labH101.Text = "00:00";
						this.labH102.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},34",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH102.Text = "LIBRE";
					}
				}
			}
			#endregion h34
			#region h35
			if (NurmdHabitacion==35)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button35.BackColor = Color.SteelBlue;
						this.labH103.BackColor = Color.SteelBlue;
						this.labH104.BackColor = Color.SteelBlue;
						this.labH105.BackColor = Color.SteelBlue;
						this.labH103.Text = Hinicia;
						this.labH104.Text = HFinaliza;
						this.labH105.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button35.BackColor = Color.Green;
						this.labH103.BackColor = Color.Green;
						this.labH104.BackColor = Color.Green;
						this.labH105.BackColor = Color.Green;
						this.labH103.Text = "00:00";
						this.labH104.Text = "00:00";
						this.labH105.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},35",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH105.Text = "LIBRE";
					}
				}
			}
			#endregion h35
			#region h36
			if (NurmdHabitacion==36)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button36.BackColor = Color.SteelBlue;
						this.labH106.BackColor = Color.SteelBlue;
						this.labH107.BackColor = Color.SteelBlue;
						this.labH108.BackColor = Color.SteelBlue;
						this.labH106.Text = Hinicia;
						this.labH107.Text = HFinaliza;
						this.labH108.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button36.BackColor = Color.Green;
						this.labH106.BackColor = Color.Green;
						this.labH107.BackColor = Color.Green;
						this.labH108.BackColor = Color.Green;
						this.labH106.Text = "00:00";
						this.labH107.Text = "00:00";
						this.labH108.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},36",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH108.Text = "LIBRE";
					}
				}
			}
			#endregion h36
			#region h37
			if (NurmdHabitacion==37)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button37.BackColor = Color.SteelBlue;
						this.labH109.BackColor = Color.SteelBlue;
						this.labH110.BackColor = Color.SteelBlue;
						this.labH111.BackColor = Color.SteelBlue;
						this.labH109.Text = Hinicia;
						this.labH110.Text = HFinaliza;
						this.labH111.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button37.BackColor = Color.Green;
						this.labH109.BackColor = Color.Green;
						this.labH110.BackColor = Color.Green;
						this.labH111.BackColor = Color.Green;
						this.labH109.Text = "00:00";
						this.labH110.Text = "00:00";
						this.labH111.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},37",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH111.Text = "LIBRE";
					}
				}
			}
			#endregion h37
			#region h38
			if (NurmdHabitacion==38)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button38.BackColor = Color.SteelBlue;
						this.labH112.BackColor = Color.SteelBlue;
						this.labH113.BackColor = Color.SteelBlue;
						this.labH114.BackColor = Color.SteelBlue;
						this.labH112.Text = Hinicia;
						this.labH113.Text = HFinaliza;
						this.labH114.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button38.BackColor = Color.Green;
						this.labH112.BackColor = Color.Green;
						this.labH113.BackColor = Color.Green;
						this.labH114.BackColor = Color.Green;
						this.labH112.Text = "00:00";
						this.labH113.Text = "00:00";
						this.labH114.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},38",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH114.Text = "LIBRE";
					}
				}
			}
			#endregion h38
			#region h39
			if (NurmdHabitacion==39)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button39.BackColor = Color.SteelBlue;
						this.labH115.BackColor = Color.SteelBlue;
						this.labH116.BackColor = Color.SteelBlue;
						this.labH117.BackColor = Color.SteelBlue;
						this.labH115.Text = Hinicia;
						this.labH116.Text = HFinaliza;
						this.labH117.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button39.BackColor = Color.Green;
						this.labH115.BackColor = Color.Green;
						this.labH116.BackColor = Color.Green;
						this.labH117.BackColor = Color.Green;
						this.labH115.Text = "00:00";
						this.labH116.Text = "00:00";
						this.labH117.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},39",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH117.Text = "LIBRE";
					}
				}
			}
			#endregion h39
			#region h40
			if (NurmdHabitacion==40)
			{
				if (e_hab==2)
				{
					if (dif_mins>0)
					{
						this.button40.BackColor = Color.SteelBlue;
						this.labH118.BackColor = Color.SteelBlue;
						this.labH119.BackColor = Color.SteelBlue;
						this.labH120.BackColor = Color.SteelBlue;
						this.labH118.Text = Hinicia;
						this.labH119.Text = HFinaliza;
						this.labH120.Text = "PASADO:"+dif_mins.ToString()+" Min.";
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
				if (e_hab==3)
				{
					if (dif_mins>=5)
					{
						this.button40.BackColor = Color.Green;
						this.labH118.BackColor = Color.Green;
						this.labH119.BackColor = Color.Green;
						this.labH120.BackColor = Color.Green;
						this.labH118.Text = "00:00";
						this.labH119.Text = "00:00";
						this.labH120.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},40",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH120.Text = "LIBRE";
					}
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
					if (dif_mins>=5)
					{
						this.button41.BackColor = Color.Green;
						this.labH121.BackColor = Color.Green;
						this.labH122.BackColor = Color.Green;
						this.labH123.BackColor = Color.Green;
						this.labH121.Text = "00:00";
						this.labH122.Text = "00:00";
						this.labH123.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},41",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH123.Text = "LIBRE";
					}
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
					if (dif_mins>=5)
					{
						this.button42.BackColor = Color.Green;
						this.labH124.BackColor = Color.Green;
						this.labH125.BackColor = Color.Green;
						this.labH126.BackColor = Color.Green;
						this.labH124.Text = "00:00";
						this.labH125.Text = "00:00";
						this.labH126.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},42",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH126.Text = "LIBRE";
					}
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
					if (dif_mins>=5)
					{
						this.button43.BackColor = Color.Green;
						this.labH127.BackColor = Color.Green;
						this.labH128.BackColor = Color.Green;
						this.labH129.BackColor = Color.Green;
						this.labH127.Text = "00:00";
						this.labH128.Text = "00:00";
						this.labH129.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},43",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH129.Text = "LIBRE";
					}
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
					if (dif_mins>=5)
					{
						this.button44.BackColor = Color.Green;
						this.labH130.BackColor = Color.Green;
						this.labH131.BackColor = Color.Green;
						this.labH132.BackColor = Color.Green;
						this.labH130.Text = "00:00";
						this.labH131.Text = "00:00";
						this.labH132.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},44",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH132.Text = "LIBRE";
					}
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
					if (dif_mins>=5)
					{
						this.button45.BackColor = Color.Green;
						this.labH133.BackColor = Color.Green;
						this.labH134.BackColor = Color.Green;
						this.labH135.BackColor = Color.Green;
						this.labH133.Text = "00:00";
						this.labH134.Text = "00:00";
						this.labH135.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},45",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH135.Text = "LIBRE";
					}
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
					if (dif_mins>=5)
					{
						this.button46.BackColor = Color.Green;
						this.labH136.BackColor = Color.Green;
						this.labH137.BackColor = Color.Green;
						this.labH138.BackColor = Color.Green;
						this.labH136.Text = "00:00";
						this.labH137.Text = "00:00";
						this.labH138.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},46",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH138.Text = "LIBRE";
					}
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
					if (dif_mins>=5)
					{
						this.button47.BackColor = Color.Green;
						this.labH139.BackColor = Color.Green;
						this.labH140.BackColor = Color.Green;
						this.labH141.BackColor = Color.Green;
						this.labH139.Text = "00:00";
						this.labH140.Text = "00:00";
						this.labH141.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},47",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH141.Text = "LIBRE";
					}
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
					if (dif_mins>=5)
					{
						this.button48.BackColor = Color.Green;
						this.labH142.BackColor = Color.Green;
						this.labH143.BackColor = Color.Green;
						this.labH144.BackColor = Color.Green;
						this.labH142.Text = "00:00";
						this.labH143.Text = "00:00";
						this.labH144.Text = "LIBRE";
						/*Actualiza habitacion a estado libre*/
//						string Qry = string.Format("Exec CambiaEdo_Mantenimiento {0},48",idBodega);
//						Funcion.EjecutaSQL(cdsSeteoF, Qry, false);
					}
					else
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
						this.labH144.Text = "LIBRE";
					}
				}
			}
			#endregion h48
			#endregion habitacion estado
		}

		static public void ColorearBotones(Form ofrm, int TotHab)		{			int i = 48;			foreach (Control oCtrl in ofrm.Controls)			{                				if (oCtrl is Button)				{					oCtrl.ForeColor = Color.White;					oCtrl.BackColor = Color.Green;					oCtrl.Text = "Habitación # " + i.ToString().Trim();					if (i>TotHab) oCtrl.Visible = false;					i --;				}			}		}
		static public void ColorearEqiquetas(Form ofrm, int TolEtiq)		{			int i = 48*3;			int te = TolEtiq*3;			int y = TolEtiq*3;			foreach (Control oCtrl in ofrm.Controls)			{                				if (oCtrl is Label)				{					oCtrl.ForeColor = Color.White;          oCtrl.BackColor = Color.Green;					if (i>te) oCtrl.Visible = false;
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

			frmFacturacionHabitaciones FH = new frmFacturacionHabitaciones(idHabitacion, iBodega, true, idCompra, bServicioPagado);
			FH.ShowDialog();

			EstadoHabitaciones();
		}

		public void FacturaHabication(int nHab,string nlaBod, int idNumroHab)
		{
			frmFCHabitaciones miFC = new frmFCHabitaciones(nHab,nlaBod,idNumroHab);
			miFC.ShowDialog();

			EstadoHabitaciones();
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(1, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=1 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(1,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}
		
		private void button2_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(2, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=2 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(2,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(3, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=3 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(3,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button4_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(4, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=4 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(4,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button5_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(5, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=5 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(5,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button6_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(6, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=6 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(6,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button7_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(7, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=7 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(7,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button8_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(8, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=8 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(8,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button9_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(9, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=9 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(9,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button10_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(10, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=10 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(10,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button11_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(11, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=11 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(11,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button12_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(12, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=12 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(12,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button13_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(13, iBodega);

//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=13 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(13,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button14_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(14, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=14 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(14,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button15_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(15, iBodega);
						
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=15 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(15,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button16_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(16, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=16 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(16,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button17_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(17, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=17 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(17,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button18_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(18, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=18 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(18,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button19_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(19, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=19 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(19,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button20_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(20, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=20 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(20,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button21_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(21, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=21 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(21,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button22_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(22, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=22 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(22,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button23_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(23, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=23 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(23,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button24_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(24, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=24 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(24,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button25_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(25, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=25 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(25,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button26_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(26, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=26 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(26,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button27_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(27, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=27 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(27,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button28_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(28, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=28 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(28,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button29_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(29, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=29 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(29,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button30_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(30, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=30 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(30,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button31_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(31, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=31 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(31,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button32_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(32, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=32 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(32,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button33_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(33, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=33 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(33,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button34_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(34, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=34 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(34,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button35_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(35, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=35 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(35,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button36_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(36, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=36 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(36,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button37_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(37, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=37 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(37,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button38_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(38, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=38 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(38,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button39_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(39, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=39 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(39,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button40_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(40, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=40 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(40,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button41_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(41, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=41 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(41,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button42_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(42, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=42 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(42,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button43_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(43, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=43 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(43,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button44_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(44, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=44 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(44,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button45_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(45, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=45 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(45,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button46_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(46, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=46 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(46,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button47_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(47, iBodega);
			
//			string Qry = "SELECT idHabitacion FROM Habitaciones WHERE NumeroHabitacion=47 and idSucursal="+idBodega.ToString();
//			idNroHab = Funcion.iEscalarSQL(cdsSeteoF, Qry);
//			FacturaHabication(47,nombrdlBodega,idNroHab);
//			EstadoHabitaciones();
		}

		private void button48_Click(object sender, System.EventArgs e)
		{
			AbreFacturaHabitacion(48, iBodega);
			
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

		private void button1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 1;
				this.MnuEstado.Show(this.button1, new Point(3,30));
			}
		}

		private void button2_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 2;
				this.MnuEstado.Show(this.button2, new Point(3,30));
			}

		}

		private void button3_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 3;
				this.MnuEstado.Show(this.button3, new Point(3,30));
			}

		}

		private void button4_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 4;
				this.MnuEstado.Show(this.button4, new Point(3,30));
			}

		}

		private void button5_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 5;
				this.MnuEstado.Show(this.button5, new Point(3,30));
			}

		}

		private void button6_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 6;
				this.MnuEstado.Show(this.button6, new Point(3,30));
			}

		}

		private void button7_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 7;
				this.MnuEstado.Show(this.button7, new Point(3,30));
			}

		}

		private void button8_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 8;
				this.MnuEstado.Show(this.button8, new Point(3,30));
			}

		}

		private void button9_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 9;
				this.MnuEstado.Show(this.button9, new Point(3,30));
			}

		}

		private void button10_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 10;
				this.MnuEstado.Show(this.button10, new Point(3,30));
			}

		}

		private void button11_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 11;
				this.MnuEstado.Show(this.button11, new Point(3,30));
			}

		}

		private void button12_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 12;
				this.MnuEstado.Show(this.button12, new Point(3,30));
			}

		}

		private void button13_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 13;
				this.MnuEstado.Show(this.button13, new Point(3,30));
			}

		}

		private void button14_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 14;
				this.MnuEstado.Show(this.button14, new Point(3,30));
			}

		}

		private void button15_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 15;
				this.MnuEstado.Show(this.button15, new Point(3,30));
			}

		}

		private void button16_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 16;
				this.MnuEstado.Show(this.button16, new Point(3,30));
			}

		}

		private void button17_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 17;
				this.MnuEstado.Show(this.button17, new Point(3,30));
			}

		}

		private void button18_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 18;
				this.MnuEstado.Show(this.button18, new Point(3,30));
			}

		}

		private void button19_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 19;
				this.MnuEstado.Show(this.button19, new Point(3,30));
			}

		}

		private void button20_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 20;
				this.MnuEstado.Show(this.button20, new Point(3,30));
			}

		}

		private void button21_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 21;
				this.MnuEstado.Show(this.button21, new Point(3,30));
			}

		}

		private void button22_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 22;
				this.MnuEstado.Show(this.button22, new Point(3,30));
			}

		}

		private void button23_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 23;
				this.MnuEstado.Show(this.button23, new Point(3,30));
			}

		}

		private void button24_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 24;
				this.MnuEstado.Show(this.button24, new Point(3,30));
			}

		}

		private void button25_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 25;
				this.MnuEstado.Show(this.button25, new Point(3,30));
			}

		}

		private void button26_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 26;
				this.MnuEstado.Show(this.button26, new Point(3,30));
			}

		}

		private void button27_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 27;
				this.MnuEstado.Show(this.button27, new Point(3,30));
			}

		}

		private void button28_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 28;
				this.MnuEstado.Show(this.button28, new Point(3,30));
			}

		}

		private void button29_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 29;
				this.MnuEstado.Show(this.button29, new Point(3,30));
			}

		}

		private void button30_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 30;
				this.MnuEstado.Show(this.button30, new Point(3,30));
			}

		}

		private void button31_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 31;
				this.MnuEstado.Show(this.button31, new Point(3,30));
			}

		}

		private void button32_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 32;
				this.MnuEstado.Show(this.button32, new Point(3,30));
			}

		}

		private void button33_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 33;
				this.MnuEstado.Show(this.button33, new Point(3,30));
			}

		}

		private void button34_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 34;
				this.MnuEstado.Show(this.button34, new Point(3,30));
			}

		}

		private void button35_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 35;
				this.MnuEstado.Show(this.button35, new Point(3,30));
			}

		}

		private void button36_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 36;
				this.MnuEstado.Show(this.button36, new Point(3,30));
			}

		}

		private void button37_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 37;
				this.MnuEstado.Show(this.button37, new Point(3,30));
			}

		}

		private void button38_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 38;
				this.MnuEstado.Show(this.button38, new Point(3,30));
			}

		}

		private void button39_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 39;
				this.MnuEstado.Show(this.button39, new Point(3,30));
			}

		}

		private void button40_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 40;
				this.MnuEstado.Show(this.button40, new Point(3,30));
			}

		}

		private void button41_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 41;
				this.MnuEstado.Show(this.button41, new Point(3,30));
			}

		}

		private void button42_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 42;
				this.MnuEstado.Show(this.button42, new Point(3,30));
			}

		}

		private void button43_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 43;
				this.MnuEstado.Show(this.button43, new Point(3,30));
			}

		}

		private void button44_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 44;
				this.MnuEstado.Show(this.button44, new Point(3,30));
			}

		}

		private void button45_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 45;
				this.MnuEstado.Show(this.button45, new Point(3,30));
			}

		}

		private void button46_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 46;
				this.MnuEstado.Show(this.button46, new Point(3,30));
			}

		}

		private void button47_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 47;
				this.MnuEstado.Show(this.button47, new Point(3,30));
			}

		}

		private void button48_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				nboton = 48;
				this.MnuEstado.Show(this.button48, new Point(3,30));
			}
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
	}
}
