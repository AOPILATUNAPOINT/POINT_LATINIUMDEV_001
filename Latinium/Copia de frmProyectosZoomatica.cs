using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
//using System.Collections.Generic;
//using System.Linq;
using System.Text;
using Microsoft.VisualBasic;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmProyectosZoomatica.
	/// </summary>
	public class frmProyectosZoomatica : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.GroupBox groupBox6;
		private System.Windows.Forms.GroupBox groupBox7;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.GroupBox grpBL8;
		private System.Windows.Forms.GroupBox grpBL1;
		private System.Windows.Forms.Label lblDay1;
		private System.Windows.Forms.Label lblDay8;
		private System.Windows.Forms.GroupBox grpBL15;
		private System.Windows.Forms.Label lblDay15;
		private System.Windows.Forms.GroupBox grpBL22;
		private System.Windows.Forms.Label lblDay22;
		private System.Windows.Forms.GroupBox grpBL2;
		private System.Windows.Forms.Label lblDay2;
		private System.Windows.Forms.GroupBox grpBL6;
		private System.Windows.Forms.Label lblDay6;
		private System.Windows.Forms.GroupBox grpBL13;
		private System.Windows.Forms.Label lblDay13;
		private System.Windows.Forms.GroupBox grpBL26;
		private System.Windows.Forms.Label lblDay26;
		private System.Windows.Forms.GroupBox grpBL19;
		private System.Windows.Forms.Label lblDay19;
		private System.Windows.Forms.GroupBox grpBL12;
		private System.Windows.Forms.Label lblDay12;
		private System.Windows.Forms.GroupBox grpBL5;
		private System.Windows.Forms.Label lblDay5;
		private System.Windows.Forms.GroupBox grpBL25;
		private System.Windows.Forms.Label lblDay25;
		private System.Windows.Forms.GroupBox grpBL18;
		private System.Windows.Forms.Label lblDay18;
		private System.Windows.Forms.GroupBox grpBL11;
		private System.Windows.Forms.Label lblDay11;
		private System.Windows.Forms.GroupBox grpBL4;
		private System.Windows.Forms.Label lblDay4;
		private System.Windows.Forms.GroupBox grpBL10;
		private System.Windows.Forms.Label lblDay10;
		private System.Windows.Forms.GroupBox grpBL17;
		private System.Windows.Forms.Label lblDay17;
		private System.Windows.Forms.GroupBox grpBL24;
		private System.Windows.Forms.Label lblDay24;
		private System.Windows.Forms.GroupBox grpBL3;
		private System.Windows.Forms.Label lblDay3;
		private System.Windows.Forms.GroupBox grpBL23;
		private System.Windows.Forms.Label lblDay23;
		private System.Windows.Forms.GroupBox grpBL16;
		private System.Windows.Forms.Label lblDay16;
		private System.Windows.Forms.GroupBox grpBL9;
		private System.Windows.Forms.Label lblDay9;
		private System.Windows.Forms.GroupBox grpBL27;
		private System.Windows.Forms.Label lblDay27;
		private System.Windows.Forms.GroupBox grpBL20;
		private System.Windows.Forms.Label lblDay20;
		private System.Windows.Forms.GroupBox grpBL28;
		private System.Windows.Forms.Label lblDay28;
		private System.Windows.Forms.GroupBox grpBL21;
		private System.Windows.Forms.Label lblDay21;
		private System.Windows.Forms.GroupBox grpBL14;
		private System.Windows.Forms.Label lblDay14;
		private System.Windows.Forms.GroupBox grpBL7;
		private System.Windows.Forms.Label lblDay7;
		private System.Windows.Forms.GroupBox grpBL29;
		private System.Windows.Forms.Label lblDay29;
		private System.Windows.Forms.GroupBox grpBL31;
		private System.Windows.Forms.Label lblDay31;
		private System.Windows.Forms.GroupBox grpBL32;
		private System.Windows.Forms.Label lblDay32;
		private System.Windows.Forms.GroupBox grpBL33;
		private System.Windows.Forms.Label lblDay33;
		private System.Windows.Forms.GroupBox grpBL34;
		private System.Windows.Forms.Label lblDay34;
		private System.Windows.Forms.GroupBox grpBL35;
		private System.Windows.Forms.Label lblDay35;
		private System.Windows.Forms.GroupBox grpBL42;
		private System.Windows.Forms.Label lblDay42;
		private System.Windows.Forms.GroupBox grpBL41;
		private System.Windows.Forms.Label lblDay41;
		private System.Windows.Forms.GroupBox grpBL40;
		private System.Windows.Forms.GroupBox grpBL39;
		private System.Windows.Forms.Label lblDay39;
		private System.Windows.Forms.GroupBox grpBL38;
		private System.Windows.Forms.Label lblDay38;
		private System.Windows.Forms.GroupBox grpBL37;
		private System.Windows.Forms.Label lblDay37;
		private System.Windows.Forms.GroupBox grpBL36;
		private System.Windows.Forms.Label lblDay36;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAño;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMes;
		private System.Windows.Forms.Label lblMes;
		private System.Windows.Forms.Button btnVer;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblDay40;
		private System.Windows.Forms.Label lblMeses;
		private System.Windows.Forms.Button btnVer17;
		private System.Windows.Forms.Label lblDay30;
		private System.Windows.Forms.GroupBox grpBL30;
		private System.Windows.Forms.Button btnVer2;
		private System.Windows.Forms.Button btnVer1;
		private System.Windows.Forms.Button btnVer3;
		private System.Windows.Forms.Button btnVer4;
		private System.Windows.Forms.Button btnVer5;
		private System.Windows.Forms.Button btnVer6;
		private System.Windows.Forms.Button btnVer7;
		private System.Windows.Forms.Button btnVer8;
		private System.Windows.Forms.Button btnVer15;
		private System.Windows.Forms.Button btnVer9;
		private System.Windows.Forms.Button btnVer10;
		private System.Windows.Forms.Button btnVer11;
		private System.Windows.Forms.Button btnVer12;
		private System.Windows.Forms.Button btnVer13;
		private System.Windows.Forms.Button btnVer14;
		private System.Windows.Forms.Button btnVer16;
		private System.Windows.Forms.Button btnVer18;
		private System.Windows.Forms.Button btnVer19;
		private System.Windows.Forms.Button btnVer20;
		private System.Windows.Forms.Button btnVer21;
		private System.Windows.Forms.Button btnVer22;
		private System.Windows.Forms.Button btnVer29;
		private System.Windows.Forms.Button btnVer36;
		private System.Windows.Forms.Button btnVer23;
		private System.Windows.Forms.Button btnVer30;
		private System.Windows.Forms.Button btnVer37;
		private System.Windows.Forms.Button btnVer24;
		private System.Windows.Forms.Button btnVer31;
		private System.Windows.Forms.Button btnVer38;
		private System.Windows.Forms.Button btnVer25;
		private System.Windows.Forms.Button btnVer26;
		private System.Windows.Forms.Button btnVer27;
		private System.Windows.Forms.Button btnVer28;
		private System.Windows.Forms.Button btnVer32;
		private System.Windows.Forms.Button btnVer33;
		private System.Windows.Forms.Button btnVer34;
		private System.Windows.Forms.Button btnVer35;
		private System.Windows.Forms.Button btnVer39;
		private System.Windows.Forms.Button btnVer40;
		private System.Windows.Forms.Button btnVer41;
		private System.Windows.Forms.Button btnVer42;
		private System.Windows.Forms.LinkLabel txtNotas1;
		private System.Windows.Forms.LinkLabel txtNotas2;
		private System.Windows.Forms.LinkLabel txtNotas3;
		private System.Windows.Forms.LinkLabel txtNotas4;
		private System.Windows.Forms.LinkLabel txtNotas5;
		private System.Windows.Forms.LinkLabel txtNotas6;
		private System.Windows.Forms.LinkLabel txtNotas7;
		private System.Windows.Forms.LinkLabel txtNotas9;
		private System.Windows.Forms.LinkLabel txtNotas8;
		private System.Windows.Forms.LinkLabel txtNotas10;
		private System.Windows.Forms.LinkLabel txtNotas11;
		private System.Windows.Forms.LinkLabel txtNotas12;
		private System.Windows.Forms.LinkLabel txtNotas13;
		private System.Windows.Forms.LinkLabel txtNotas14;
		private System.Windows.Forms.LinkLabel txtNotas16;
		private System.Windows.Forms.LinkLabel txtNotas23;
		private System.Windows.Forms.LinkLabel txtNotas30;
		private System.Windows.Forms.LinkLabel txtNotas37;
		private System.Windows.Forms.LinkLabel txtNotas15;
		private System.Windows.Forms.LinkLabel txtNotas17;
		private System.Windows.Forms.LinkLabel txtNotas18;
		private System.Windows.Forms.LinkLabel txtNotas19;
		private System.Windows.Forms.LinkLabel txtNotas20;
		private System.Windows.Forms.LinkLabel txtNotas21;
		private System.Windows.Forms.LinkLabel txtNotas22;
		private System.Windows.Forms.LinkLabel txtNotas24;
		private System.Windows.Forms.LinkLabel txtNotas25;
		private System.Windows.Forms.LinkLabel txtNotas26;
		private System.Windows.Forms.LinkLabel txtNotas27;
		private System.Windows.Forms.LinkLabel txtNotas28;
		private System.Windows.Forms.LinkLabel txtNotas29;
		private System.Windows.Forms.LinkLabel txtNotas31;
		private System.Windows.Forms.LinkLabel txtNotas32;
		private System.Windows.Forms.LinkLabel txtNotas33;
		private System.Windows.Forms.LinkLabel txtNotas34;
		private System.Windows.Forms.LinkLabel txtNotas35;
		private System.Windows.Forms.LinkLabel txtNotas36;
		private System.Windows.Forms.LinkLabel txtNotas38;
		private System.Windows.Forms.LinkLabel txtNotas39;
		private System.Windows.Forms.LinkLabel txtNotas40;
		private System.Windows.Forms.LinkLabel txtNotas41;
		private System.Windows.Forms.LinkLabel txtNotas42;
		private System.Windows.Forms.LinkLabel linkLabel1;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.ToolTip toolTip2;
		private System.Windows.Forms.ToolTip toolTip3;
		private System.ComponentModel.IContainer components;

		public frmProyectosZoomatica()
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
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmProyectosZoomatica));
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label4 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.label3 = new System.Windows.Forms.Label();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.label5 = new System.Windows.Forms.Label();
			this.groupBox6 = new System.Windows.Forms.GroupBox();
			this.label6 = new System.Windows.Forms.Label();
			this.groupBox7 = new System.Windows.Forms.GroupBox();
			this.label7 = new System.Windows.Forms.Label();
			this.grpBL1 = new System.Windows.Forms.GroupBox();
			this.btnVer1 = new System.Windows.Forms.Button();
			this.lblDay1 = new System.Windows.Forms.Label();
			this.txtNotas1 = new System.Windows.Forms.LinkLabel();
			this.grpBL8 = new System.Windows.Forms.GroupBox();
			this.txtNotas8 = new System.Windows.Forms.LinkLabel();
			this.btnVer8 = new System.Windows.Forms.Button();
			this.lblDay8 = new System.Windows.Forms.Label();
			this.grpBL15 = new System.Windows.Forms.GroupBox();
			this.txtNotas15 = new System.Windows.Forms.LinkLabel();
			this.btnVer15 = new System.Windows.Forms.Button();
			this.lblDay15 = new System.Windows.Forms.Label();
			this.grpBL22 = new System.Windows.Forms.GroupBox();
			this.txtNotas22 = new System.Windows.Forms.LinkLabel();
			this.btnVer22 = new System.Windows.Forms.Button();
			this.lblDay22 = new System.Windows.Forms.Label();
			this.grpBL2 = new System.Windows.Forms.GroupBox();
			this.txtNotas2 = new System.Windows.Forms.LinkLabel();
			this.lblDay2 = new System.Windows.Forms.Label();
			this.btnVer2 = new System.Windows.Forms.Button();
			this.grpBL6 = new System.Windows.Forms.GroupBox();
			this.txtNotas6 = new System.Windows.Forms.LinkLabel();
			this.btnVer6 = new System.Windows.Forms.Button();
			this.lblDay6 = new System.Windows.Forms.Label();
			this.grpBL13 = new System.Windows.Forms.GroupBox();
			this.txtNotas13 = new System.Windows.Forms.LinkLabel();
			this.btnVer13 = new System.Windows.Forms.Button();
			this.lblDay13 = new System.Windows.Forms.Label();
			this.grpBL26 = new System.Windows.Forms.GroupBox();
			this.txtNotas26 = new System.Windows.Forms.LinkLabel();
			this.btnVer26 = new System.Windows.Forms.Button();
			this.lblDay26 = new System.Windows.Forms.Label();
			this.grpBL19 = new System.Windows.Forms.GroupBox();
			this.txtNotas19 = new System.Windows.Forms.LinkLabel();
			this.btnVer19 = new System.Windows.Forms.Button();
			this.lblDay19 = new System.Windows.Forms.Label();
			this.grpBL12 = new System.Windows.Forms.GroupBox();
			this.txtNotas12 = new System.Windows.Forms.LinkLabel();
			this.btnVer12 = new System.Windows.Forms.Button();
			this.lblDay12 = new System.Windows.Forms.Label();
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			this.grpBL5 = new System.Windows.Forms.GroupBox();
			this.txtNotas5 = new System.Windows.Forms.LinkLabel();
			this.btnVer5 = new System.Windows.Forms.Button();
			this.lblDay5 = new System.Windows.Forms.Label();
			this.grpBL25 = new System.Windows.Forms.GroupBox();
			this.txtNotas25 = new System.Windows.Forms.LinkLabel();
			this.btnVer25 = new System.Windows.Forms.Button();
			this.lblDay25 = new System.Windows.Forms.Label();
			this.grpBL18 = new System.Windows.Forms.GroupBox();
			this.txtNotas18 = new System.Windows.Forms.LinkLabel();
			this.btnVer18 = new System.Windows.Forms.Button();
			this.lblDay18 = new System.Windows.Forms.Label();
			this.grpBL11 = new System.Windows.Forms.GroupBox();
			this.txtNotas11 = new System.Windows.Forms.LinkLabel();
			this.btnVer11 = new System.Windows.Forms.Button();
			this.lblDay11 = new System.Windows.Forms.Label();
			this.grpBL4 = new System.Windows.Forms.GroupBox();
			this.txtNotas4 = new System.Windows.Forms.LinkLabel();
			this.btnVer4 = new System.Windows.Forms.Button();
			this.lblDay4 = new System.Windows.Forms.Label();
			this.grpBL10 = new System.Windows.Forms.GroupBox();
			this.txtNotas10 = new System.Windows.Forms.LinkLabel();
			this.btnVer10 = new System.Windows.Forms.Button();
			this.lblDay10 = new System.Windows.Forms.Label();
			this.grpBL17 = new System.Windows.Forms.GroupBox();
			this.txtNotas17 = new System.Windows.Forms.LinkLabel();
			this.btnVer17 = new System.Windows.Forms.Button();
			this.lblDay17 = new System.Windows.Forms.Label();
			this.grpBL24 = new System.Windows.Forms.GroupBox();
			this.txtNotas24 = new System.Windows.Forms.LinkLabel();
			this.btnVer24 = new System.Windows.Forms.Button();
			this.lblDay24 = new System.Windows.Forms.Label();
			this.grpBL3 = new System.Windows.Forms.GroupBox();
			this.txtNotas3 = new System.Windows.Forms.LinkLabel();
			this.btnVer3 = new System.Windows.Forms.Button();
			this.lblDay3 = new System.Windows.Forms.Label();
			this.grpBL23 = new System.Windows.Forms.GroupBox();
			this.txtNotas23 = new System.Windows.Forms.LinkLabel();
			this.btnVer23 = new System.Windows.Forms.Button();
			this.lblDay23 = new System.Windows.Forms.Label();
			this.grpBL16 = new System.Windows.Forms.GroupBox();
			this.txtNotas16 = new System.Windows.Forms.LinkLabel();
			this.btnVer16 = new System.Windows.Forms.Button();
			this.lblDay16 = new System.Windows.Forms.Label();
			this.grpBL9 = new System.Windows.Forms.GroupBox();
			this.txtNotas9 = new System.Windows.Forms.LinkLabel();
			this.btnVer9 = new System.Windows.Forms.Button();
			this.lblDay9 = new System.Windows.Forms.Label();
			this.grpBL27 = new System.Windows.Forms.GroupBox();
			this.txtNotas27 = new System.Windows.Forms.LinkLabel();
			this.btnVer27 = new System.Windows.Forms.Button();
			this.lblDay27 = new System.Windows.Forms.Label();
			this.grpBL20 = new System.Windows.Forms.GroupBox();
			this.txtNotas20 = new System.Windows.Forms.LinkLabel();
			this.btnVer20 = new System.Windows.Forms.Button();
			this.lblDay20 = new System.Windows.Forms.Label();
			this.grpBL28 = new System.Windows.Forms.GroupBox();
			this.txtNotas28 = new System.Windows.Forms.LinkLabel();
			this.btnVer28 = new System.Windows.Forms.Button();
			this.lblDay28 = new System.Windows.Forms.Label();
			this.grpBL21 = new System.Windows.Forms.GroupBox();
			this.txtNotas21 = new System.Windows.Forms.LinkLabel();
			this.btnVer21 = new System.Windows.Forms.Button();
			this.lblDay21 = new System.Windows.Forms.Label();
			this.grpBL14 = new System.Windows.Forms.GroupBox();
			this.txtNotas14 = new System.Windows.Forms.LinkLabel();
			this.btnVer14 = new System.Windows.Forms.Button();
			this.lblDay14 = new System.Windows.Forms.Label();
			this.grpBL7 = new System.Windows.Forms.GroupBox();
			this.txtNotas7 = new System.Windows.Forms.LinkLabel();
			this.btnVer7 = new System.Windows.Forms.Button();
			this.lblDay7 = new System.Windows.Forms.Label();
			this.grpBL29 = new System.Windows.Forms.GroupBox();
			this.txtNotas29 = new System.Windows.Forms.LinkLabel();
			this.btnVer29 = new System.Windows.Forms.Button();
			this.lblDay29 = new System.Windows.Forms.Label();
			this.grpBL31 = new System.Windows.Forms.GroupBox();
			this.txtNotas31 = new System.Windows.Forms.LinkLabel();
			this.btnVer31 = new System.Windows.Forms.Button();
			this.lblDay31 = new System.Windows.Forms.Label();
			this.grpBL32 = new System.Windows.Forms.GroupBox();
			this.txtNotas32 = new System.Windows.Forms.LinkLabel();
			this.btnVer32 = new System.Windows.Forms.Button();
			this.lblDay32 = new System.Windows.Forms.Label();
			this.grpBL33 = new System.Windows.Forms.GroupBox();
			this.txtNotas33 = new System.Windows.Forms.LinkLabel();
			this.btnVer33 = new System.Windows.Forms.Button();
			this.lblDay33 = new System.Windows.Forms.Label();
			this.grpBL34 = new System.Windows.Forms.GroupBox();
			this.txtNotas34 = new System.Windows.Forms.LinkLabel();
			this.btnVer34 = new System.Windows.Forms.Button();
			this.lblDay34 = new System.Windows.Forms.Label();
			this.grpBL35 = new System.Windows.Forms.GroupBox();
			this.txtNotas35 = new System.Windows.Forms.LinkLabel();
			this.btnVer35 = new System.Windows.Forms.Button();
			this.lblDay35 = new System.Windows.Forms.Label();
			this.grpBL42 = new System.Windows.Forms.GroupBox();
			this.txtNotas42 = new System.Windows.Forms.LinkLabel();
			this.btnVer42 = new System.Windows.Forms.Button();
			this.lblDay42 = new System.Windows.Forms.Label();
			this.grpBL41 = new System.Windows.Forms.GroupBox();
			this.txtNotas41 = new System.Windows.Forms.LinkLabel();
			this.btnVer41 = new System.Windows.Forms.Button();
			this.lblDay41 = new System.Windows.Forms.Label();
			this.grpBL40 = new System.Windows.Forms.GroupBox();
			this.txtNotas40 = new System.Windows.Forms.LinkLabel();
			this.btnVer40 = new System.Windows.Forms.Button();
			this.lblDay40 = new System.Windows.Forms.Label();
			this.grpBL39 = new System.Windows.Forms.GroupBox();
			this.txtNotas39 = new System.Windows.Forms.LinkLabel();
			this.btnVer39 = new System.Windows.Forms.Button();
			this.lblDay39 = new System.Windows.Forms.Label();
			this.grpBL38 = new System.Windows.Forms.GroupBox();
			this.txtNotas38 = new System.Windows.Forms.LinkLabel();
			this.btnVer38 = new System.Windows.Forms.Button();
			this.lblDay38 = new System.Windows.Forms.Label();
			this.grpBL37 = new System.Windows.Forms.GroupBox();
			this.txtNotas37 = new System.Windows.Forms.LinkLabel();
			this.btnVer37 = new System.Windows.Forms.Button();
			this.lblDay37 = new System.Windows.Forms.Label();
			this.grpBL36 = new System.Windows.Forms.GroupBox();
			this.txtNotas36 = new System.Windows.Forms.LinkLabel();
			this.btnVer36 = new System.Windows.Forms.Button();
			this.lblDay36 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.txtAño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtMes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblMes = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblMeses = new System.Windows.Forms.Label();
			this.lblDay30 = new System.Windows.Forms.Label();
			this.grpBL30 = new System.Windows.Forms.GroupBox();
			this.txtNotas30 = new System.Windows.Forms.LinkLabel();
			this.btnVer30 = new System.Windows.Forms.Button();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.button1 = new System.Windows.Forms.Button();
			this.toolTip2 = new System.Windows.Forms.ToolTip(this.components);
			this.toolTip3 = new System.Windows.Forms.ToolTip(this.components);
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.groupBox4.SuspendLayout();
			this.groupBox5.SuspendLayout();
			this.groupBox6.SuspendLayout();
			this.groupBox7.SuspendLayout();
			this.grpBL1.SuspendLayout();
			this.grpBL8.SuspendLayout();
			this.grpBL15.SuspendLayout();
			this.grpBL22.SuspendLayout();
			this.grpBL2.SuspendLayout();
			this.grpBL6.SuspendLayout();
			this.grpBL13.SuspendLayout();
			this.grpBL26.SuspendLayout();
			this.grpBL19.SuspendLayout();
			this.grpBL12.SuspendLayout();
			this.grpBL5.SuspendLayout();
			this.grpBL25.SuspendLayout();
			this.grpBL18.SuspendLayout();
			this.grpBL11.SuspendLayout();
			this.grpBL4.SuspendLayout();
			this.grpBL10.SuspendLayout();
			this.grpBL17.SuspendLayout();
			this.grpBL24.SuspendLayout();
			this.grpBL3.SuspendLayout();
			this.grpBL23.SuspendLayout();
			this.grpBL16.SuspendLayout();
			this.grpBL9.SuspendLayout();
			this.grpBL27.SuspendLayout();
			this.grpBL20.SuspendLayout();
			this.grpBL28.SuspendLayout();
			this.grpBL21.SuspendLayout();
			this.grpBL14.SuspendLayout();
			this.grpBL7.SuspendLayout();
			this.grpBL29.SuspendLayout();
			this.grpBL31.SuspendLayout();
			this.grpBL32.SuspendLayout();
			this.grpBL33.SuspendLayout();
			this.grpBL34.SuspendLayout();
			this.grpBL35.SuspendLayout();
			this.grpBL42.SuspendLayout();
			this.grpBL41.SuspendLayout();
			this.grpBL40.SuspendLayout();
			this.grpBL39.SuspendLayout();
			this.grpBL38.SuspendLayout();
			this.grpBL37.SuspendLayout();
			this.grpBL36.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.grpBL30.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(16, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(184, 32);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(53, 12);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.label4.TabIndex = 61;
			this.label4.Text = "LUNES";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(200, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(184, 32);
			this.groupBox2.TabIndex = 1;
			this.groupBox2.TabStop = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(55, 12);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(53, 16);
			this.label1.TabIndex = 62;
			this.label1.Text = "MARTES";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.label2);
			this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox3.Location = new System.Drawing.Point(384, 40);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(184, 32);
			this.groupBox3.TabIndex = 2;
			this.groupBox3.TabStop = false;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(55, 12);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(73, 16);
			this.label2.TabIndex = 62;
			this.label2.Text = "MIÉRCOLES";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox4
			// 
			this.groupBox4.Controls.Add(this.label3);
			this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox4.Location = new System.Drawing.Point(568, 40);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(184, 32);
			this.groupBox4.TabIndex = 3;
			this.groupBox4.TabStop = false;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(55, 12);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(50, 16);
			this.label3.TabIndex = 62;
			this.label3.Text = "JUEVES";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox5
			// 
			this.groupBox5.Controls.Add(this.label5);
			this.groupBox5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox5.Location = new System.Drawing.Point(752, 40);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(184, 32);
			this.groupBox5.TabIndex = 4;
			this.groupBox5.TabStop = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(55, 12);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(55, 16);
			this.label5.TabIndex = 62;
			this.label5.Text = "VIERNES";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox6
			// 
			this.groupBox6.Controls.Add(this.label6);
			this.groupBox6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox6.Location = new System.Drawing.Point(936, 40);
			this.groupBox6.Name = "groupBox6";
			this.groupBox6.Size = new System.Drawing.Size(184, 32);
			this.groupBox6.TabIndex = 5;
			this.groupBox6.TabStop = false;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(55, 12);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(53, 16);
			this.label6.TabIndex = 62;
			this.label6.Text = "SÁBADO";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox7
			// 
			this.groupBox7.Controls.Add(this.label7);
			this.groupBox7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox7.Location = new System.Drawing.Point(1120, 40);
			this.groupBox7.Name = "groupBox7";
			this.groupBox7.Size = new System.Drawing.Size(184, 32);
			this.groupBox7.TabIndex = 6;
			this.groupBox7.TabStop = false;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(55, 12);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(61, 16);
			this.label7.TabIndex = 62;
			this.label7.Text = "DOMINGO";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL1
			// 
			this.grpBL1.BackColor = System.Drawing.SystemColors.Control;
			this.grpBL1.Controls.Add(this.btnVer1);
			this.grpBL1.Controls.Add(this.lblDay1);
			this.grpBL1.Controls.Add(this.txtNotas1);
			this.grpBL1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL1.Location = new System.Drawing.Point(16, 72);
			this.grpBL1.Name = "grpBL1";
			this.grpBL1.Size = new System.Drawing.Size(184, 112);
			this.grpBL1.TabIndex = 7;
			this.grpBL1.TabStop = false;
			// 
			// btnVer1
			// 
			this.btnVer1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer1.Location = new System.Drawing.Point(8, 8);
			this.btnVer1.Name = "btnVer1";
			this.btnVer1.Size = new System.Drawing.Size(168, 8);
			this.btnVer1.TabIndex = 761;
			this.toolTip1.SetToolTip(this.btnVer1, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay1
			// 
			this.lblDay1.AutoSize = true;
			this.lblDay1.BackColor = System.Drawing.Color.Transparent;
			this.lblDay1.Location = new System.Drawing.Point(8, 16);
			this.lblDay1.Name = "lblDay1";
			this.lblDay1.Size = new System.Drawing.Size(0, 16);
			this.lblDay1.TabIndex = 61;
			this.lblDay1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNotas1
			// 
			this.txtNotas1.LinkColor = System.Drawing.Color.Black;
			this.txtNotas1.Location = new System.Drawing.Point(8, 16);
			this.txtNotas1.Name = "txtNotas1";
			this.txtNotas1.Size = new System.Drawing.Size(172, 80);
			this.txtNotas1.TabIndex = 760;
			// 
			// grpBL8
			// 
			this.grpBL8.Controls.Add(this.txtNotas8);
			this.grpBL8.Controls.Add(this.btnVer8);
			this.grpBL8.Controls.Add(this.lblDay8);
			this.grpBL8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL8.Location = new System.Drawing.Point(16, 184);
			this.grpBL8.Name = "grpBL8";
			this.grpBL8.Size = new System.Drawing.Size(184, 104);
			this.grpBL8.TabIndex = 8;
			this.grpBL8.TabStop = false;
			// 
			// txtNotas8
			// 
			this.txtNotas8.LinkColor = System.Drawing.Color.Black;
			this.txtNotas8.Location = new System.Drawing.Point(6, 13);
			this.txtNotas8.Name = "txtNotas8";
			this.txtNotas8.Size = new System.Drawing.Size(172, 79);
			this.txtNotas8.TabIndex = 764;
			// 
			// btnVer8
			// 
			this.btnVer8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer8.Location = new System.Drawing.Point(8, 92);
			this.btnVer8.Name = "btnVer8";
			this.btnVer8.Size = new System.Drawing.Size(168, 8);
			this.btnVer8.TabIndex = 762;
			this.toolTip1.SetToolTip(this.btnVer8, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay8
			// 
			this.lblDay8.AutoSize = true;
			this.lblDay8.BackColor = System.Drawing.Color.Transparent;
			this.lblDay8.Location = new System.Drawing.Point(16, 16);
			this.lblDay8.Name = "lblDay8";
			this.lblDay8.Size = new System.Drawing.Size(0, 16);
			this.lblDay8.TabIndex = 61;
			this.lblDay8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL15
			// 
			this.grpBL15.Controls.Add(this.txtNotas15);
			this.grpBL15.Controls.Add(this.btnVer15);
			this.grpBL15.Controls.Add(this.lblDay15);
			this.grpBL15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL15.Location = new System.Drawing.Point(16, 288);
			this.grpBL15.Name = "grpBL15";
			this.grpBL15.Size = new System.Drawing.Size(184, 104);
			this.grpBL15.TabIndex = 9;
			this.grpBL15.TabStop = false;
			// 
			// txtNotas15
			// 
			this.txtNotas15.LinkColor = System.Drawing.Color.Black;
			this.txtNotas15.Location = new System.Drawing.Point(6, 13);
			this.txtNotas15.Name = "txtNotas15";
			this.txtNotas15.Size = new System.Drawing.Size(172, 78);
			this.txtNotas15.TabIndex = 765;
			// 
			// btnVer15
			// 
			this.btnVer15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer15.Location = new System.Drawing.Point(8, 90);
			this.btnVer15.Name = "btnVer15";
			this.btnVer15.Size = new System.Drawing.Size(168, 8);
			this.btnVer15.TabIndex = 762;
			this.toolTip1.SetToolTip(this.btnVer15, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay15
			// 
			this.lblDay15.AutoSize = true;
			this.lblDay15.BackColor = System.Drawing.Color.Transparent;
			this.lblDay15.Location = new System.Drawing.Point(16, 16);
			this.lblDay15.Name = "lblDay15";
			this.lblDay15.Size = new System.Drawing.Size(0, 16);
			this.lblDay15.TabIndex = 61;
			this.lblDay15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL22
			// 
			this.grpBL22.Controls.Add(this.txtNotas22);
			this.grpBL22.Controls.Add(this.btnVer22);
			this.grpBL22.Controls.Add(this.lblDay22);
			this.grpBL22.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL22.Location = new System.Drawing.Point(16, 392);
			this.grpBL22.Name = "grpBL22";
			this.grpBL22.Size = new System.Drawing.Size(184, 96);
			this.grpBL22.TabIndex = 10;
			this.grpBL22.TabStop = false;
			// 
			// txtNotas22
			// 
			this.txtNotas22.LinkColor = System.Drawing.Color.Black;
			this.txtNotas22.Location = new System.Drawing.Point(6, 14);
			this.txtNotas22.Name = "txtNotas22";
			this.txtNotas22.Size = new System.Drawing.Size(172, 68);
			this.txtNotas22.TabIndex = 766;
			// 
			// btnVer22
			// 
			this.btnVer22.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer22.Location = new System.Drawing.Point(8, 82);
			this.btnVer22.Name = "btnVer22";
			this.btnVer22.Size = new System.Drawing.Size(168, 8);
			this.btnVer22.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer22, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay22
			// 
			this.lblDay22.AutoSize = true;
			this.lblDay22.BackColor = System.Drawing.Color.Transparent;
			this.lblDay22.Location = new System.Drawing.Point(16, 16);
			this.lblDay22.Name = "lblDay22";
			this.lblDay22.Size = new System.Drawing.Size(0, 16);
			this.lblDay22.TabIndex = 61;
			this.lblDay22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL2
			// 
			this.grpBL2.Controls.Add(this.txtNotas2);
			this.grpBL2.Controls.Add(this.lblDay2);
			this.grpBL2.Controls.Add(this.btnVer2);
			this.grpBL2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL2.Location = new System.Drawing.Point(200, 72);
			this.grpBL2.Name = "grpBL2";
			this.grpBL2.Size = new System.Drawing.Size(184, 112);
			this.grpBL2.TabIndex = 11;
			this.grpBL2.TabStop = false;
			// 
			// txtNotas2
			// 
			this.txtNotas2.LinkColor = System.Drawing.Color.Black;
			this.txtNotas2.Location = new System.Drawing.Point(6, 15);
			this.txtNotas2.Name = "txtNotas2";
			this.txtNotas2.Size = new System.Drawing.Size(172, 82);
			this.txtNotas2.TabIndex = 761;
			// 
			// lblDay2
			// 
			this.lblDay2.AutoSize = true;
			this.lblDay2.BackColor = System.Drawing.Color.Transparent;
			this.lblDay2.Location = new System.Drawing.Point(8, 16);
			this.lblDay2.Name = "lblDay2";
			this.lblDay2.Size = new System.Drawing.Size(0, 16);
			this.lblDay2.TabIndex = 61;
			this.lblDay2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer2
			// 
			this.btnVer2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer2.Location = new System.Drawing.Point(8, 97);
			this.btnVer2.Name = "btnVer2";
			this.btnVer2.Size = new System.Drawing.Size(168, 8);
			this.btnVer2.TabIndex = 760;
			this.toolTip1.SetToolTip(this.btnVer2, "Click Para Visualizar Proyectos del Día");
			// 
			// grpBL6
			// 
			this.grpBL6.Controls.Add(this.txtNotas6);
			this.grpBL6.Controls.Add(this.btnVer6);
			this.grpBL6.Controls.Add(this.lblDay6);
			this.grpBL6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL6.Location = new System.Drawing.Point(936, 72);
			this.grpBL6.Name = "grpBL6";
			this.grpBL6.Size = new System.Drawing.Size(184, 112);
			this.grpBL6.TabIndex = 12;
			this.grpBL6.TabStop = false;
			// 
			// txtNotas6
			// 
			this.txtNotas6.LinkColor = System.Drawing.Color.Black;
			this.txtNotas6.Location = new System.Drawing.Point(6, 15);
			this.txtNotas6.Name = "txtNotas6";
			this.txtNotas6.Size = new System.Drawing.Size(172, 82);
			this.txtNotas6.TabIndex = 762;
			// 
			// btnVer6
			// 
			this.btnVer6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer6.Location = new System.Drawing.Point(8, 98);
			this.btnVer6.Name = "btnVer6";
			this.btnVer6.Size = new System.Drawing.Size(168, 8);
			this.btnVer6.TabIndex = 761;
			this.toolTip1.SetToolTip(this.btnVer6, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay6
			// 
			this.lblDay6.AutoSize = true;
			this.lblDay6.BackColor = System.Drawing.Color.Transparent;
			this.lblDay6.Location = new System.Drawing.Point(8, 16);
			this.lblDay6.Name = "lblDay6";
			this.lblDay6.Size = new System.Drawing.Size(0, 16);
			this.lblDay6.TabIndex = 61;
			this.lblDay6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL13
			// 
			this.grpBL13.Controls.Add(this.txtNotas13);
			this.grpBL13.Controls.Add(this.btnVer13);
			this.grpBL13.Controls.Add(this.lblDay13);
			this.grpBL13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL13.Location = new System.Drawing.Point(936, 184);
			this.grpBL13.Name = "grpBL13";
			this.grpBL13.Size = new System.Drawing.Size(184, 104);
			this.grpBL13.TabIndex = 13;
			this.grpBL13.TabStop = false;
			// 
			// txtNotas13
			// 
			this.txtNotas13.LinkColor = System.Drawing.Color.Black;
			this.txtNotas13.Location = new System.Drawing.Point(6, 13);
			this.txtNotas13.Name = "txtNotas13";
			this.txtNotas13.Size = new System.Drawing.Size(172, 79);
			this.txtNotas13.TabIndex = 764;
			this.txtNotas13.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel12_LinkClicked);
			// 
			// btnVer13
			// 
			this.btnVer13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer13.Location = new System.Drawing.Point(8, 92);
			this.btnVer13.Name = "btnVer13";
			this.btnVer13.Size = new System.Drawing.Size(168, 8);
			this.btnVer13.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer13, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay13
			// 
			this.lblDay13.AutoSize = true;
			this.lblDay13.BackColor = System.Drawing.Color.Transparent;
			this.lblDay13.Location = new System.Drawing.Point(8, 16);
			this.lblDay13.Name = "lblDay13";
			this.lblDay13.Size = new System.Drawing.Size(0, 16);
			this.lblDay13.TabIndex = 61;
			this.lblDay13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL26
			// 
			this.grpBL26.Controls.Add(this.txtNotas26);
			this.grpBL26.Controls.Add(this.btnVer26);
			this.grpBL26.Controls.Add(this.lblDay26);
			this.grpBL26.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL26.Location = new System.Drawing.Point(752, 392);
			this.grpBL26.Name = "grpBL26";
			this.grpBL26.Size = new System.Drawing.Size(184, 96);
			this.grpBL26.TabIndex = 14;
			this.grpBL26.TabStop = false;
			// 
			// txtNotas26
			// 
			this.txtNotas26.LinkColor = System.Drawing.Color.Black;
			this.txtNotas26.Location = new System.Drawing.Point(6, 14);
			this.txtNotas26.Name = "txtNotas26";
			this.txtNotas26.Size = new System.Drawing.Size(172, 68);
			this.txtNotas26.TabIndex = 766;
			// 
			// btnVer26
			// 
			this.btnVer26.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer26.Location = new System.Drawing.Point(8, 82);
			this.btnVer26.Name = "btnVer26";
			this.btnVer26.Size = new System.Drawing.Size(168, 8);
			this.btnVer26.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer26, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay26
			// 
			this.lblDay26.AutoSize = true;
			this.lblDay26.BackColor = System.Drawing.Color.Transparent;
			this.lblDay26.Location = new System.Drawing.Point(8, 16);
			this.lblDay26.Name = "lblDay26";
			this.lblDay26.Size = new System.Drawing.Size(0, 16);
			this.lblDay26.TabIndex = 61;
			this.lblDay26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL19
			// 
			this.grpBL19.Controls.Add(this.txtNotas19);
			this.grpBL19.Controls.Add(this.btnVer19);
			this.grpBL19.Controls.Add(this.lblDay19);
			this.grpBL19.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL19.Location = new System.Drawing.Point(752, 288);
			this.grpBL19.Name = "grpBL19";
			this.grpBL19.Size = new System.Drawing.Size(184, 104);
			this.grpBL19.TabIndex = 15;
			this.grpBL19.TabStop = false;
			// 
			// txtNotas19
			// 
			this.txtNotas19.LinkColor = System.Drawing.Color.Black;
			this.txtNotas19.Location = new System.Drawing.Point(6, 13);
			this.txtNotas19.Name = "txtNotas19";
			this.txtNotas19.Size = new System.Drawing.Size(172, 78);
			this.txtNotas19.TabIndex = 765;
			// 
			// btnVer19
			// 
			this.btnVer19.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer19.Location = new System.Drawing.Point(8, 90);
			this.btnVer19.Name = "btnVer19";
			this.btnVer19.Size = new System.Drawing.Size(168, 8);
			this.btnVer19.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer19, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay19
			// 
			this.lblDay19.AutoSize = true;
			this.lblDay19.BackColor = System.Drawing.Color.Transparent;
			this.lblDay19.Location = new System.Drawing.Point(8, 16);
			this.lblDay19.Name = "lblDay19";
			this.lblDay19.Size = new System.Drawing.Size(0, 16);
			this.lblDay19.TabIndex = 61;
			this.lblDay19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL12
			// 
			this.grpBL12.Controls.Add(this.txtNotas12);
			this.grpBL12.Controls.Add(this.btnVer12);
			this.grpBL12.Controls.Add(this.lblDay12);
			this.grpBL12.Controls.Add(this.linkLabel1);
			this.grpBL12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL12.Location = new System.Drawing.Point(752, 184);
			this.grpBL12.Name = "grpBL12";
			this.grpBL12.Size = new System.Drawing.Size(184, 104);
			this.grpBL12.TabIndex = 16;
			this.grpBL12.TabStop = false;
			// 
			// txtNotas12
			// 
			this.txtNotas12.LinkColor = System.Drawing.Color.Black;
			this.txtNotas12.Location = new System.Drawing.Point(6, 13);
			this.txtNotas12.Name = "txtNotas12";
			this.txtNotas12.Size = new System.Drawing.Size(172, 79);
			this.txtNotas12.TabIndex = 764;
			// 
			// btnVer12
			// 
			this.btnVer12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer12.Location = new System.Drawing.Point(8, 92);
			this.btnVer12.Name = "btnVer12";
			this.btnVer12.Size = new System.Drawing.Size(168, 8);
			this.btnVer12.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer12, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay12
			// 
			this.lblDay12.AutoSize = true;
			this.lblDay12.BackColor = System.Drawing.Color.Transparent;
			this.lblDay12.Location = new System.Drawing.Point(8, 16);
			this.lblDay12.Name = "lblDay12";
			this.lblDay12.Size = new System.Drawing.Size(0, 16);
			this.lblDay12.TabIndex = 61;
			this.lblDay12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// linkLabel1
			// 
			this.linkLabel1.Location = new System.Drawing.Point(0, 8);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(172, 79);
			this.linkLabel1.TabIndex = 764;
			// 
			// grpBL5
			// 
			this.grpBL5.Controls.Add(this.txtNotas5);
			this.grpBL5.Controls.Add(this.btnVer5);
			this.grpBL5.Controls.Add(this.lblDay5);
			this.grpBL5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL5.Location = new System.Drawing.Point(752, 72);
			this.grpBL5.Name = "grpBL5";
			this.grpBL5.Size = new System.Drawing.Size(184, 112);
			this.grpBL5.TabIndex = 17;
			this.grpBL5.TabStop = false;
			// 
			// txtNotas5
			// 
			this.txtNotas5.LinkColor = System.Drawing.Color.Black;
			this.txtNotas5.Location = new System.Drawing.Point(6, 15);
			this.txtNotas5.Name = "txtNotas5";
			this.txtNotas5.Size = new System.Drawing.Size(172, 82);
			this.txtNotas5.TabIndex = 762;
			// 
			// btnVer5
			// 
			this.btnVer5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer5.Location = new System.Drawing.Point(8, 98);
			this.btnVer5.Name = "btnVer5";
			this.btnVer5.Size = new System.Drawing.Size(168, 8);
			this.btnVer5.TabIndex = 761;
			this.toolTip1.SetToolTip(this.btnVer5, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay5
			// 
			this.lblDay5.AutoSize = true;
			this.lblDay5.BackColor = System.Drawing.Color.Transparent;
			this.lblDay5.Location = new System.Drawing.Point(8, 16);
			this.lblDay5.Name = "lblDay5";
			this.lblDay5.Size = new System.Drawing.Size(0, 16);
			this.lblDay5.TabIndex = 61;
			this.lblDay5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL25
			// 
			this.grpBL25.Controls.Add(this.txtNotas25);
			this.grpBL25.Controls.Add(this.btnVer25);
			this.grpBL25.Controls.Add(this.lblDay25);
			this.grpBL25.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL25.Location = new System.Drawing.Point(568, 392);
			this.grpBL25.Name = "grpBL25";
			this.grpBL25.Size = new System.Drawing.Size(184, 96);
			this.grpBL25.TabIndex = 18;
			this.grpBL25.TabStop = false;
			// 
			// txtNotas25
			// 
			this.txtNotas25.LinkColor = System.Drawing.Color.Black;
			this.txtNotas25.Location = new System.Drawing.Point(6, 14);
			this.txtNotas25.Name = "txtNotas25";
			this.txtNotas25.Size = new System.Drawing.Size(172, 68);
			this.txtNotas25.TabIndex = 766;
			// 
			// btnVer25
			// 
			this.btnVer25.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer25.Location = new System.Drawing.Point(8, 82);
			this.btnVer25.Name = "btnVer25";
			this.btnVer25.Size = new System.Drawing.Size(168, 8);
			this.btnVer25.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer25, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay25
			// 
			this.lblDay25.AutoSize = true;
			this.lblDay25.BackColor = System.Drawing.Color.Transparent;
			this.lblDay25.Location = new System.Drawing.Point(8, 16);
			this.lblDay25.Name = "lblDay25";
			this.lblDay25.Size = new System.Drawing.Size(0, 16);
			this.lblDay25.TabIndex = 61;
			this.lblDay25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL18
			// 
			this.grpBL18.Controls.Add(this.txtNotas18);
			this.grpBL18.Controls.Add(this.btnVer18);
			this.grpBL18.Controls.Add(this.lblDay18);
			this.grpBL18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL18.Location = new System.Drawing.Point(568, 288);
			this.grpBL18.Name = "grpBL18";
			this.grpBL18.Size = new System.Drawing.Size(184, 104);
			this.grpBL18.TabIndex = 19;
			this.grpBL18.TabStop = false;
			// 
			// txtNotas18
			// 
			this.txtNotas18.LinkColor = System.Drawing.Color.Black;
			this.txtNotas18.Location = new System.Drawing.Point(6, 80);
			this.txtNotas18.Name = "txtNotas18";
			this.txtNotas18.Size = new System.Drawing.Size(172, 8);
			this.txtNotas18.TabIndex = 765;
			// 
			// btnVer18
			// 
			this.btnVer18.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer18.Location = new System.Drawing.Point(8, 90);
			this.btnVer18.Name = "btnVer18";
			this.btnVer18.Size = new System.Drawing.Size(168, 8);
			this.btnVer18.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer18, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay18
			// 
			this.lblDay18.AutoSize = true;
			this.lblDay18.BackColor = System.Drawing.Color.Transparent;
			this.lblDay18.Location = new System.Drawing.Point(8, 16);
			this.lblDay18.Name = "lblDay18";
			this.lblDay18.Size = new System.Drawing.Size(0, 16);
			this.lblDay18.TabIndex = 61;
			this.lblDay18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL11
			// 
			this.grpBL11.Controls.Add(this.txtNotas11);
			this.grpBL11.Controls.Add(this.btnVer11);
			this.grpBL11.Controls.Add(this.lblDay11);
			this.grpBL11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL11.Location = new System.Drawing.Point(568, 184);
			this.grpBL11.Name = "grpBL11";
			this.grpBL11.Size = new System.Drawing.Size(184, 104);
			this.grpBL11.TabIndex = 20;
			this.grpBL11.TabStop = false;
			// 
			// txtNotas11
			// 
			this.txtNotas11.LinkColor = System.Drawing.Color.Black;
			this.txtNotas11.Location = new System.Drawing.Point(6, 13);
			this.txtNotas11.Name = "txtNotas11";
			this.txtNotas11.Size = new System.Drawing.Size(172, 79);
			this.txtNotas11.TabIndex = 764;
			// 
			// btnVer11
			// 
			this.btnVer11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer11.Location = new System.Drawing.Point(8, 92);
			this.btnVer11.Name = "btnVer11";
			this.btnVer11.Size = new System.Drawing.Size(168, 8);
			this.btnVer11.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer11, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay11
			// 
			this.lblDay11.AutoSize = true;
			this.lblDay11.BackColor = System.Drawing.Color.Transparent;
			this.lblDay11.Location = new System.Drawing.Point(8, 16);
			this.lblDay11.Name = "lblDay11";
			this.lblDay11.Size = new System.Drawing.Size(0, 16);
			this.lblDay11.TabIndex = 61;
			this.lblDay11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL4
			// 
			this.grpBL4.Controls.Add(this.txtNotas4);
			this.grpBL4.Controls.Add(this.btnVer4);
			this.grpBL4.Controls.Add(this.lblDay4);
			this.grpBL4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL4.Location = new System.Drawing.Point(568, 72);
			this.grpBL4.Name = "grpBL4";
			this.grpBL4.Size = new System.Drawing.Size(184, 112);
			this.grpBL4.TabIndex = 21;
			this.grpBL4.TabStop = false;
			// 
			// txtNotas4
			// 
			this.txtNotas4.LinkColor = System.Drawing.Color.Black;
			this.txtNotas4.Location = new System.Drawing.Point(6, 15);
			this.txtNotas4.Name = "txtNotas4";
			this.txtNotas4.Size = new System.Drawing.Size(172, 82);
			this.txtNotas4.TabIndex = 762;
			// 
			// btnVer4
			// 
			this.btnVer4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer4.Location = new System.Drawing.Point(8, 98);
			this.btnVer4.Name = "btnVer4";
			this.btnVer4.Size = new System.Drawing.Size(168, 8);
			this.btnVer4.TabIndex = 761;
			this.toolTip1.SetToolTip(this.btnVer4, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay4
			// 
			this.lblDay4.AutoSize = true;
			this.lblDay4.BackColor = System.Drawing.Color.Transparent;
			this.lblDay4.Location = new System.Drawing.Point(8, 16);
			this.lblDay4.Name = "lblDay4";
			this.lblDay4.Size = new System.Drawing.Size(0, 16);
			this.lblDay4.TabIndex = 61;
			this.lblDay4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL10
			// 
			this.grpBL10.Controls.Add(this.txtNotas10);
			this.grpBL10.Controls.Add(this.btnVer10);
			this.grpBL10.Controls.Add(this.lblDay10);
			this.grpBL10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL10.Location = new System.Drawing.Point(384, 184);
			this.grpBL10.Name = "grpBL10";
			this.grpBL10.Size = new System.Drawing.Size(184, 104);
			this.grpBL10.TabIndex = 22;
			this.grpBL10.TabStop = false;
			// 
			// txtNotas10
			// 
			this.txtNotas10.LinkColor = System.Drawing.Color.Black;
			this.txtNotas10.Location = new System.Drawing.Point(6, 13);
			this.txtNotas10.Name = "txtNotas10";
			this.txtNotas10.Size = new System.Drawing.Size(172, 79);
			this.txtNotas10.TabIndex = 764;
			// 
			// btnVer10
			// 
			this.btnVer10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer10.Location = new System.Drawing.Point(8, 92);
			this.btnVer10.Name = "btnVer10";
			this.btnVer10.Size = new System.Drawing.Size(168, 8);
			this.btnVer10.TabIndex = 762;
			this.toolTip1.SetToolTip(this.btnVer10, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay10
			// 
			this.lblDay10.AutoSize = true;
			this.lblDay10.BackColor = System.Drawing.Color.Transparent;
			this.lblDay10.Location = new System.Drawing.Point(8, 16);
			this.lblDay10.Name = "lblDay10";
			this.lblDay10.Size = new System.Drawing.Size(0, 16);
			this.lblDay10.TabIndex = 61;
			this.lblDay10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL17
			// 
			this.grpBL17.Controls.Add(this.txtNotas17);
			this.grpBL17.Controls.Add(this.btnVer17);
			this.grpBL17.Controls.Add(this.lblDay17);
			this.grpBL17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL17.Location = new System.Drawing.Point(384, 288);
			this.grpBL17.Name = "grpBL17";
			this.grpBL17.Size = new System.Drawing.Size(184, 104);
			this.grpBL17.TabIndex = 23;
			this.grpBL17.TabStop = false;
			// 
			// txtNotas17
			// 
			this.txtNotas17.LinkColor = System.Drawing.Color.Black;
			this.txtNotas17.Location = new System.Drawing.Point(6, 80);
			this.txtNotas17.Name = "txtNotas17";
			this.txtNotas17.Size = new System.Drawing.Size(172, 8);
			this.txtNotas17.TabIndex = 765;
			// 
			// btnVer17
			// 
			this.btnVer17.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer17.Location = new System.Drawing.Point(8, 90);
			this.btnVer17.Name = "btnVer17";
			this.btnVer17.Size = new System.Drawing.Size(168, 8);
			this.btnVer17.TabIndex = 761;
			this.toolTip1.SetToolTip(this.btnVer17, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay17
			// 
			this.lblDay17.AutoSize = true;
			this.lblDay17.BackColor = System.Drawing.Color.Transparent;
			this.lblDay17.Location = new System.Drawing.Point(8, 16);
			this.lblDay17.Name = "lblDay17";
			this.lblDay17.Size = new System.Drawing.Size(0, 16);
			this.lblDay17.TabIndex = 61;
			this.lblDay17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL24
			// 
			this.grpBL24.Controls.Add(this.txtNotas24);
			this.grpBL24.Controls.Add(this.btnVer24);
			this.grpBL24.Controls.Add(this.lblDay24);
			this.grpBL24.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL24.Location = new System.Drawing.Point(384, 392);
			this.grpBL24.Name = "grpBL24";
			this.grpBL24.Size = new System.Drawing.Size(184, 96);
			this.grpBL24.TabIndex = 24;
			this.grpBL24.TabStop = false;
			// 
			// txtNotas24
			// 
			this.txtNotas24.LinkColor = System.Drawing.Color.Black;
			this.txtNotas24.Location = new System.Drawing.Point(6, 14);
			this.txtNotas24.Name = "txtNotas24";
			this.txtNotas24.Size = new System.Drawing.Size(172, 68);
			this.txtNotas24.TabIndex = 766;
			// 
			// btnVer24
			// 
			this.btnVer24.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer24.Location = new System.Drawing.Point(8, 82);
			this.btnVer24.Name = "btnVer24";
			this.btnVer24.Size = new System.Drawing.Size(168, 8);
			this.btnVer24.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer24, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay24
			// 
			this.lblDay24.AutoSize = true;
			this.lblDay24.BackColor = System.Drawing.Color.Transparent;
			this.lblDay24.Location = new System.Drawing.Point(8, 16);
			this.lblDay24.Name = "lblDay24";
			this.lblDay24.Size = new System.Drawing.Size(0, 16);
			this.lblDay24.TabIndex = 61;
			this.lblDay24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL3
			// 
			this.grpBL3.Controls.Add(this.txtNotas3);
			this.grpBL3.Controls.Add(this.btnVer3);
			this.grpBL3.Controls.Add(this.lblDay3);
			this.grpBL3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL3.Location = new System.Drawing.Point(384, 72);
			this.grpBL3.Name = "grpBL3";
			this.grpBL3.Size = new System.Drawing.Size(184, 112);
			this.grpBL3.TabIndex = 25;
			this.grpBL3.TabStop = false;
			// 
			// txtNotas3
			// 
			this.txtNotas3.LinkColor = System.Drawing.Color.Black;
			this.txtNotas3.Location = new System.Drawing.Point(6, 15);
			this.txtNotas3.Name = "txtNotas3";
			this.txtNotas3.Size = new System.Drawing.Size(172, 82);
			this.txtNotas3.TabIndex = 762;
			// 
			// btnVer3
			// 
			this.btnVer3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer3.Location = new System.Drawing.Point(8, 98);
			this.btnVer3.Name = "btnVer3";
			this.btnVer3.Size = new System.Drawing.Size(168, 8);
			this.btnVer3.TabIndex = 761;
			this.toolTip1.SetToolTip(this.btnVer3, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay3
			// 
			this.lblDay3.AutoSize = true;
			this.lblDay3.BackColor = System.Drawing.Color.Transparent;
			this.lblDay3.Location = new System.Drawing.Point(8, 16);
			this.lblDay3.Name = "lblDay3";
			this.lblDay3.Size = new System.Drawing.Size(0, 16);
			this.lblDay3.TabIndex = 61;
			this.lblDay3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL23
			// 
			this.grpBL23.Controls.Add(this.txtNotas23);
			this.grpBL23.Controls.Add(this.btnVer23);
			this.grpBL23.Controls.Add(this.lblDay23);
			this.grpBL23.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL23.Location = new System.Drawing.Point(200, 392);
			this.grpBL23.Name = "grpBL23";
			this.grpBL23.Size = new System.Drawing.Size(184, 96);
			this.grpBL23.TabIndex = 26;
			this.grpBL23.TabStop = false;
			// 
			// txtNotas23
			// 
			this.txtNotas23.LinkColor = System.Drawing.Color.Black;
			this.txtNotas23.Location = new System.Drawing.Point(6, 13);
			this.txtNotas23.Name = "txtNotas23";
			this.txtNotas23.Size = new System.Drawing.Size(172, 68);
			this.txtNotas23.TabIndex = 765;
			// 
			// btnVer23
			// 
			this.btnVer23.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer23.Location = new System.Drawing.Point(8, 82);
			this.btnVer23.Name = "btnVer23";
			this.btnVer23.Size = new System.Drawing.Size(168, 8);
			this.btnVer23.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer23, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay23
			// 
			this.lblDay23.AutoSize = true;
			this.lblDay23.BackColor = System.Drawing.Color.Transparent;
			this.lblDay23.Location = new System.Drawing.Point(16, 16);
			this.lblDay23.Name = "lblDay23";
			this.lblDay23.Size = new System.Drawing.Size(0, 16);
			this.lblDay23.TabIndex = 61;
			this.lblDay23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL16
			// 
			this.grpBL16.Controls.Add(this.txtNotas16);
			this.grpBL16.Controls.Add(this.btnVer16);
			this.grpBL16.Controls.Add(this.lblDay16);
			this.grpBL16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL16.Location = new System.Drawing.Point(200, 288);
			this.grpBL16.Name = "grpBL16";
			this.grpBL16.Size = new System.Drawing.Size(184, 104);
			this.grpBL16.TabIndex = 27;
			this.grpBL16.TabStop = false;
			// 
			// txtNotas16
			// 
			this.txtNotas16.LinkColor = System.Drawing.Color.Black;
			this.txtNotas16.Location = new System.Drawing.Point(6, 11);
			this.txtNotas16.Name = "txtNotas16";
			this.txtNotas16.Size = new System.Drawing.Size(172, 78);
			this.txtNotas16.TabIndex = 764;
			// 
			// btnVer16
			// 
			this.btnVer16.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer16.Location = new System.Drawing.Point(8, 90);
			this.btnVer16.Name = "btnVer16";
			this.btnVer16.Size = new System.Drawing.Size(168, 8);
			this.btnVer16.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer16, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay16
			// 
			this.lblDay16.AutoSize = true;
			this.lblDay16.BackColor = System.Drawing.Color.Transparent;
			this.lblDay16.Location = new System.Drawing.Point(8, 16);
			this.lblDay16.Name = "lblDay16";
			this.lblDay16.Size = new System.Drawing.Size(0, 16);
			this.lblDay16.TabIndex = 61;
			this.lblDay16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblDay16.Click += new System.EventHandler(this.label28_Click);
			// 
			// grpBL9
			// 
			this.grpBL9.Controls.Add(this.txtNotas9);
			this.grpBL9.Controls.Add(this.btnVer9);
			this.grpBL9.Controls.Add(this.lblDay9);
			this.grpBL9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL9.Location = new System.Drawing.Point(200, 184);
			this.grpBL9.Name = "grpBL9";
			this.grpBL9.Size = new System.Drawing.Size(184, 104);
			this.grpBL9.TabIndex = 28;
			this.grpBL9.TabStop = false;
			// 
			// txtNotas9
			// 
			this.txtNotas9.LinkColor = System.Drawing.Color.Black;
			this.txtNotas9.Location = new System.Drawing.Point(6, 11);
			this.txtNotas9.Name = "txtNotas9";
			this.txtNotas9.Size = new System.Drawing.Size(172, 79);
			this.txtNotas9.TabIndex = 763;
			// 
			// btnVer9
			// 
			this.btnVer9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer9.Location = new System.Drawing.Point(8, 92);
			this.btnVer9.Name = "btnVer9";
			this.btnVer9.Size = new System.Drawing.Size(168, 8);
			this.btnVer9.TabIndex = 762;
			this.toolTip1.SetToolTip(this.btnVer9, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay9
			// 
			this.lblDay9.AutoSize = true;
			this.lblDay9.BackColor = System.Drawing.Color.Transparent;
			this.lblDay9.Location = new System.Drawing.Point(8, 16);
			this.lblDay9.Name = "lblDay9";
			this.lblDay9.Size = new System.Drawing.Size(0, 16);
			this.lblDay9.TabIndex = 61;
			this.lblDay9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL27
			// 
			this.grpBL27.Controls.Add(this.txtNotas27);
			this.grpBL27.Controls.Add(this.btnVer27);
			this.grpBL27.Controls.Add(this.lblDay27);
			this.grpBL27.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL27.Location = new System.Drawing.Point(936, 392);
			this.grpBL27.Name = "grpBL27";
			this.grpBL27.Size = new System.Drawing.Size(184, 96);
			this.grpBL27.TabIndex = 30;
			this.grpBL27.TabStop = false;
			// 
			// txtNotas27
			// 
			this.txtNotas27.LinkColor = System.Drawing.Color.Black;
			this.txtNotas27.Location = new System.Drawing.Point(6, 14);
			this.txtNotas27.Name = "txtNotas27";
			this.txtNotas27.Size = new System.Drawing.Size(172, 68);
			this.txtNotas27.TabIndex = 766;
			// 
			// btnVer27
			// 
			this.btnVer27.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer27.Location = new System.Drawing.Point(8, 82);
			this.btnVer27.Name = "btnVer27";
			this.btnVer27.Size = new System.Drawing.Size(168, 8);
			this.btnVer27.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer27, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay27
			// 
			this.lblDay27.AutoSize = true;
			this.lblDay27.BackColor = System.Drawing.Color.Transparent;
			this.lblDay27.Location = new System.Drawing.Point(8, 16);
			this.lblDay27.Name = "lblDay27";
			this.lblDay27.Size = new System.Drawing.Size(0, 16);
			this.lblDay27.TabIndex = 61;
			this.lblDay27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL20
			// 
			this.grpBL20.Controls.Add(this.txtNotas20);
			this.grpBL20.Controls.Add(this.btnVer20);
			this.grpBL20.Controls.Add(this.lblDay20);
			this.grpBL20.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL20.Location = new System.Drawing.Point(936, 288);
			this.grpBL20.Name = "grpBL20";
			this.grpBL20.Size = new System.Drawing.Size(184, 104);
			this.grpBL20.TabIndex = 29;
			this.grpBL20.TabStop = false;
			// 
			// txtNotas20
			// 
			this.txtNotas20.LinkColor = System.Drawing.Color.Black;
			this.txtNotas20.Location = new System.Drawing.Point(6, 13);
			this.txtNotas20.Name = "txtNotas20";
			this.txtNotas20.Size = new System.Drawing.Size(172, 78);
			this.txtNotas20.TabIndex = 765;
			// 
			// btnVer20
			// 
			this.btnVer20.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer20.Location = new System.Drawing.Point(8, 90);
			this.btnVer20.Name = "btnVer20";
			this.btnVer20.Size = new System.Drawing.Size(168, 8);
			this.btnVer20.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer20, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay20
			// 
			this.lblDay20.AutoSize = true;
			this.lblDay20.BackColor = System.Drawing.Color.Transparent;
			this.lblDay20.Location = new System.Drawing.Point(8, 16);
			this.lblDay20.Name = "lblDay20";
			this.lblDay20.Size = new System.Drawing.Size(0, 16);
			this.lblDay20.TabIndex = 61;
			this.lblDay20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL28
			// 
			this.grpBL28.Controls.Add(this.txtNotas28);
			this.grpBL28.Controls.Add(this.btnVer28);
			this.grpBL28.Controls.Add(this.lblDay28);
			this.grpBL28.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL28.Location = new System.Drawing.Point(1120, 392);
			this.grpBL28.Name = "grpBL28";
			this.grpBL28.Size = new System.Drawing.Size(184, 96);
			this.grpBL28.TabIndex = 34;
			this.grpBL28.TabStop = false;
			// 
			// txtNotas28
			// 
			this.txtNotas28.LinkColor = System.Drawing.Color.Black;
			this.txtNotas28.Location = new System.Drawing.Point(6, 14);
			this.txtNotas28.Name = "txtNotas28";
			this.txtNotas28.Size = new System.Drawing.Size(172, 68);
			this.txtNotas28.TabIndex = 766;
			// 
			// btnVer28
			// 
			this.btnVer28.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer28.Location = new System.Drawing.Point(8, 82);
			this.btnVer28.Name = "btnVer28";
			this.btnVer28.Size = new System.Drawing.Size(168, 8);
			this.btnVer28.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer28, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay28
			// 
			this.lblDay28.AutoSize = true;
			this.lblDay28.BackColor = System.Drawing.Color.Transparent;
			this.lblDay28.Location = new System.Drawing.Point(8, 16);
			this.lblDay28.Name = "lblDay28";
			this.lblDay28.Size = new System.Drawing.Size(0, 16);
			this.lblDay28.TabIndex = 61;
			this.lblDay28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL21
			// 
			this.grpBL21.Controls.Add(this.txtNotas21);
			this.grpBL21.Controls.Add(this.btnVer21);
			this.grpBL21.Controls.Add(this.lblDay21);
			this.grpBL21.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL21.Location = new System.Drawing.Point(1120, 288);
			this.grpBL21.Name = "grpBL21";
			this.grpBL21.Size = new System.Drawing.Size(184, 104);
			this.grpBL21.TabIndex = 33;
			this.grpBL21.TabStop = false;
			// 
			// txtNotas21
			// 
			this.txtNotas21.LinkColor = System.Drawing.Color.Black;
			this.txtNotas21.Location = new System.Drawing.Point(6, 13);
			this.txtNotas21.Name = "txtNotas21";
			this.txtNotas21.Size = new System.Drawing.Size(172, 78);
			this.txtNotas21.TabIndex = 765;
			// 
			// btnVer21
			// 
			this.btnVer21.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer21.Location = new System.Drawing.Point(8, 90);
			this.btnVer21.Name = "btnVer21";
			this.btnVer21.Size = new System.Drawing.Size(168, 8);
			this.btnVer21.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer21, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay21
			// 
			this.lblDay21.AutoSize = true;
			this.lblDay21.BackColor = System.Drawing.Color.Transparent;
			this.lblDay21.Location = new System.Drawing.Point(8, 16);
			this.lblDay21.Name = "lblDay21";
			this.lblDay21.Size = new System.Drawing.Size(0, 16);
			this.lblDay21.TabIndex = 61;
			this.lblDay21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL14
			// 
			this.grpBL14.Controls.Add(this.txtNotas14);
			this.grpBL14.Controls.Add(this.btnVer14);
			this.grpBL14.Controls.Add(this.lblDay14);
			this.grpBL14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL14.Location = new System.Drawing.Point(1120, 184);
			this.grpBL14.Name = "grpBL14";
			this.grpBL14.Size = new System.Drawing.Size(184, 104);
			this.grpBL14.TabIndex = 32;
			this.grpBL14.TabStop = false;
			// 
			// txtNotas14
			// 
			this.txtNotas14.LinkColor = System.Drawing.Color.Black;
			this.txtNotas14.Location = new System.Drawing.Point(6, 13);
			this.txtNotas14.Name = "txtNotas14";
			this.txtNotas14.Size = new System.Drawing.Size(172, 79);
			this.txtNotas14.TabIndex = 764;
			// 
			// btnVer14
			// 
			this.btnVer14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer14.Location = new System.Drawing.Point(8, 92);
			this.btnVer14.Name = "btnVer14";
			this.btnVer14.Size = new System.Drawing.Size(168, 8);
			this.btnVer14.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer14, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay14
			// 
			this.lblDay14.AutoSize = true;
			this.lblDay14.BackColor = System.Drawing.Color.Transparent;
			this.lblDay14.Location = new System.Drawing.Point(8, 16);
			this.lblDay14.Name = "lblDay14";
			this.lblDay14.Size = new System.Drawing.Size(0, 16);
			this.lblDay14.TabIndex = 61;
			this.lblDay14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL7
			// 
			this.grpBL7.Controls.Add(this.txtNotas7);
			this.grpBL7.Controls.Add(this.btnVer7);
			this.grpBL7.Controls.Add(this.lblDay7);
			this.grpBL7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL7.Location = new System.Drawing.Point(1120, 72);
			this.grpBL7.Name = "grpBL7";
			this.grpBL7.Size = new System.Drawing.Size(184, 112);
			this.grpBL7.TabIndex = 31;
			this.grpBL7.TabStop = false;
			// 
			// txtNotas7
			// 
			this.txtNotas7.LinkColor = System.Drawing.Color.Black;
			this.txtNotas7.Location = new System.Drawing.Point(6, 15);
			this.txtNotas7.Name = "txtNotas7";
			this.txtNotas7.Size = new System.Drawing.Size(172, 82);
			this.txtNotas7.TabIndex = 762;
			// 
			// btnVer7
			// 
			this.btnVer7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer7.Location = new System.Drawing.Point(8, 98);
			this.btnVer7.Name = "btnVer7";
			this.btnVer7.Size = new System.Drawing.Size(168, 8);
			this.btnVer7.TabIndex = 761;
			this.toolTip1.SetToolTip(this.btnVer7, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay7
			// 
			this.lblDay7.AutoSize = true;
			this.lblDay7.BackColor = System.Drawing.Color.Transparent;
			this.lblDay7.Location = new System.Drawing.Point(8, 16);
			this.lblDay7.Name = "lblDay7";
			this.lblDay7.Size = new System.Drawing.Size(0, 16);
			this.lblDay7.TabIndex = 61;
			this.lblDay7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL29
			// 
			this.grpBL29.Controls.Add(this.txtNotas29);
			this.grpBL29.Controls.Add(this.btnVer29);
			this.grpBL29.Controls.Add(this.lblDay29);
			this.grpBL29.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL29.Location = new System.Drawing.Point(16, 488);
			this.grpBL29.Name = "grpBL29";
			this.grpBL29.Size = new System.Drawing.Size(184, 104);
			this.grpBL29.TabIndex = 35;
			this.grpBL29.TabStop = false;
			// 
			// txtNotas29
			// 
			this.txtNotas29.LinkColor = System.Drawing.Color.Black;
			this.txtNotas29.Location = new System.Drawing.Point(6, 12);
			this.txtNotas29.Name = "txtNotas29";
			this.txtNotas29.Size = new System.Drawing.Size(172, 80);
			this.txtNotas29.TabIndex = 767;
			// 
			// btnVer29
			// 
			this.btnVer29.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer29.Location = new System.Drawing.Point(8, 92);
			this.btnVer29.Name = "btnVer29";
			this.btnVer29.Size = new System.Drawing.Size(168, 8);
			this.btnVer29.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer29, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay29
			// 
			this.lblDay29.AutoSize = true;
			this.lblDay29.BackColor = System.Drawing.Color.Transparent;
			this.lblDay29.Location = new System.Drawing.Point(16, 16);
			this.lblDay29.Name = "lblDay29";
			this.lblDay29.Size = new System.Drawing.Size(0, 16);
			this.lblDay29.TabIndex = 61;
			this.lblDay29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL31
			// 
			this.grpBL31.Controls.Add(this.txtNotas31);
			this.grpBL31.Controls.Add(this.btnVer31);
			this.grpBL31.Controls.Add(this.lblDay31);
			this.grpBL31.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL31.Location = new System.Drawing.Point(384, 488);
			this.grpBL31.Name = "grpBL31";
			this.grpBL31.Size = new System.Drawing.Size(184, 104);
			this.grpBL31.TabIndex = 37;
			this.grpBL31.TabStop = false;
			// 
			// txtNotas31
			// 
			this.txtNotas31.LinkColor = System.Drawing.Color.Black;
			this.txtNotas31.Location = new System.Drawing.Point(6, 12);
			this.txtNotas31.Name = "txtNotas31";
			this.txtNotas31.Size = new System.Drawing.Size(172, 80);
			this.txtNotas31.TabIndex = 767;
			// 
			// btnVer31
			// 
			this.btnVer31.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer31.Location = new System.Drawing.Point(8, 92);
			this.btnVer31.Name = "btnVer31";
			this.btnVer31.Size = new System.Drawing.Size(168, 8);
			this.btnVer31.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer31, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay31
			// 
			this.lblDay31.AutoSize = true;
			this.lblDay31.BackColor = System.Drawing.Color.Transparent;
			this.lblDay31.Location = new System.Drawing.Point(16, 16);
			this.lblDay31.Name = "lblDay31";
			this.lblDay31.Size = new System.Drawing.Size(0, 16);
			this.lblDay31.TabIndex = 61;
			this.lblDay31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL32
			// 
			this.grpBL32.Controls.Add(this.txtNotas32);
			this.grpBL32.Controls.Add(this.btnVer32);
			this.grpBL32.Controls.Add(this.lblDay32);
			this.grpBL32.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL32.Location = new System.Drawing.Point(568, 488);
			this.grpBL32.Name = "grpBL32";
			this.grpBL32.Size = new System.Drawing.Size(184, 104);
			this.grpBL32.TabIndex = 38;
			this.grpBL32.TabStop = false;
			// 
			// txtNotas32
			// 
			this.txtNotas32.LinkColor = System.Drawing.Color.Black;
			this.txtNotas32.Location = new System.Drawing.Point(6, 12);
			this.txtNotas32.Name = "txtNotas32";
			this.txtNotas32.Size = new System.Drawing.Size(172, 80);
			this.txtNotas32.TabIndex = 767;
			// 
			// btnVer32
			// 
			this.btnVer32.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer32.Location = new System.Drawing.Point(8, 92);
			this.btnVer32.Name = "btnVer32";
			this.btnVer32.Size = new System.Drawing.Size(168, 8);
			this.btnVer32.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer32, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay32
			// 
			this.lblDay32.AutoSize = true;
			this.lblDay32.BackColor = System.Drawing.Color.Transparent;
			this.lblDay32.Location = new System.Drawing.Point(16, 16);
			this.lblDay32.Name = "lblDay32";
			this.lblDay32.Size = new System.Drawing.Size(0, 16);
			this.lblDay32.TabIndex = 61;
			this.lblDay32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL33
			// 
			this.grpBL33.Controls.Add(this.txtNotas33);
			this.grpBL33.Controls.Add(this.btnVer33);
			this.grpBL33.Controls.Add(this.lblDay33);
			this.grpBL33.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL33.Location = new System.Drawing.Point(752, 488);
			this.grpBL33.Name = "grpBL33";
			this.grpBL33.Size = new System.Drawing.Size(184, 104);
			this.grpBL33.TabIndex = 39;
			this.grpBL33.TabStop = false;
			// 
			// txtNotas33
			// 
			this.txtNotas33.LinkColor = System.Drawing.Color.Black;
			this.txtNotas33.Location = new System.Drawing.Point(6, 12);
			this.txtNotas33.Name = "txtNotas33";
			this.txtNotas33.Size = new System.Drawing.Size(172, 80);
			this.txtNotas33.TabIndex = 767;
			// 
			// btnVer33
			// 
			this.btnVer33.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer33.Location = new System.Drawing.Point(8, 92);
			this.btnVer33.Name = "btnVer33";
			this.btnVer33.Size = new System.Drawing.Size(168, 8);
			this.btnVer33.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer33, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay33
			// 
			this.lblDay33.AutoSize = true;
			this.lblDay33.BackColor = System.Drawing.Color.Transparent;
			this.lblDay33.Location = new System.Drawing.Point(16, 16);
			this.lblDay33.Name = "lblDay33";
			this.lblDay33.Size = new System.Drawing.Size(0, 16);
			this.lblDay33.TabIndex = 61;
			this.lblDay33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL34
			// 
			this.grpBL34.Controls.Add(this.txtNotas34);
			this.grpBL34.Controls.Add(this.btnVer34);
			this.grpBL34.Controls.Add(this.lblDay34);
			this.grpBL34.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL34.Location = new System.Drawing.Point(936, 488);
			this.grpBL34.Name = "grpBL34";
			this.grpBL34.Size = new System.Drawing.Size(184, 104);
			this.grpBL34.TabIndex = 40;
			this.grpBL34.TabStop = false;
			// 
			// txtNotas34
			// 
			this.txtNotas34.LinkColor = System.Drawing.Color.Black;
			this.txtNotas34.Location = new System.Drawing.Point(6, 12);
			this.txtNotas34.Name = "txtNotas34";
			this.txtNotas34.Size = new System.Drawing.Size(172, 80);
			this.txtNotas34.TabIndex = 767;
			// 
			// btnVer34
			// 
			this.btnVer34.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer34.Location = new System.Drawing.Point(8, 92);
			this.btnVer34.Name = "btnVer34";
			this.btnVer34.Size = new System.Drawing.Size(168, 8);
			this.btnVer34.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer34, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay34
			// 
			this.lblDay34.AutoSize = true;
			this.lblDay34.BackColor = System.Drawing.Color.Transparent;
			this.lblDay34.Location = new System.Drawing.Point(16, 16);
			this.lblDay34.Name = "lblDay34";
			this.lblDay34.Size = new System.Drawing.Size(0, 16);
			this.lblDay34.TabIndex = 61;
			this.lblDay34.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL35
			// 
			this.grpBL35.Controls.Add(this.txtNotas35);
			this.grpBL35.Controls.Add(this.btnVer35);
			this.grpBL35.Controls.Add(this.lblDay35);
			this.grpBL35.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL35.Location = new System.Drawing.Point(1120, 488);
			this.grpBL35.Name = "grpBL35";
			this.grpBL35.Size = new System.Drawing.Size(184, 104);
			this.grpBL35.TabIndex = 41;
			this.grpBL35.TabStop = false;
			// 
			// txtNotas35
			// 
			this.txtNotas35.LinkColor = System.Drawing.Color.Black;
			this.txtNotas35.Location = new System.Drawing.Point(6, 12);
			this.txtNotas35.Name = "txtNotas35";
			this.txtNotas35.Size = new System.Drawing.Size(172, 80);
			this.txtNotas35.TabIndex = 767;
			// 
			// btnVer35
			// 
			this.btnVer35.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer35.Location = new System.Drawing.Point(8, 92);
			this.btnVer35.Name = "btnVer35";
			this.btnVer35.Size = new System.Drawing.Size(168, 8);
			this.btnVer35.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer35, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay35
			// 
			this.lblDay35.AutoSize = true;
			this.lblDay35.BackColor = System.Drawing.Color.Transparent;
			this.lblDay35.Location = new System.Drawing.Point(16, 16);
			this.lblDay35.Name = "lblDay35";
			this.lblDay35.Size = new System.Drawing.Size(0, 16);
			this.lblDay35.TabIndex = 61;
			this.lblDay35.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL42
			// 
			this.grpBL42.Controls.Add(this.txtNotas42);
			this.grpBL42.Controls.Add(this.btnVer42);
			this.grpBL42.Controls.Add(this.lblDay42);
			this.grpBL42.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL42.Location = new System.Drawing.Point(1120, 592);
			this.grpBL42.Name = "grpBL42";
			this.grpBL42.Size = new System.Drawing.Size(184, 112);
			this.grpBL42.TabIndex = 48;
			this.grpBL42.TabStop = false;
			// 
			// txtNotas42
			// 
			this.txtNotas42.LinkColor = System.Drawing.Color.Black;
			this.txtNotas42.Location = new System.Drawing.Point(6, 15);
			this.txtNotas42.Name = "txtNotas42";
			this.txtNotas42.Size = new System.Drawing.Size(172, 82);
			this.txtNotas42.TabIndex = 767;
			// 
			// btnVer42
			// 
			this.btnVer42.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer42.Location = new System.Drawing.Point(8, 98);
			this.btnVer42.Name = "btnVer42";
			this.btnVer42.Size = new System.Drawing.Size(168, 8);
			this.btnVer42.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer42, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay42
			// 
			this.lblDay42.AutoSize = true;
			this.lblDay42.BackColor = System.Drawing.Color.Transparent;
			this.lblDay42.Location = new System.Drawing.Point(16, 16);
			this.lblDay42.Name = "lblDay42";
			this.lblDay42.Size = new System.Drawing.Size(0, 16);
			this.lblDay42.TabIndex = 61;
			this.lblDay42.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL41
			// 
			this.grpBL41.Controls.Add(this.txtNotas41);
			this.grpBL41.Controls.Add(this.btnVer41);
			this.grpBL41.Controls.Add(this.lblDay41);
			this.grpBL41.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL41.Location = new System.Drawing.Point(936, 592);
			this.grpBL41.Name = "grpBL41";
			this.grpBL41.Size = new System.Drawing.Size(184, 112);
			this.grpBL41.TabIndex = 47;
			this.grpBL41.TabStop = false;
			// 
			// txtNotas41
			// 
			this.txtNotas41.LinkColor = System.Drawing.Color.Black;
			this.txtNotas41.Location = new System.Drawing.Point(6, 15);
			this.txtNotas41.Name = "txtNotas41";
			this.txtNotas41.Size = new System.Drawing.Size(172, 82);
			this.txtNotas41.TabIndex = 767;
			// 
			// btnVer41
			// 
			this.btnVer41.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer41.Location = new System.Drawing.Point(8, 98);
			this.btnVer41.Name = "btnVer41";
			this.btnVer41.Size = new System.Drawing.Size(168, 8);
			this.btnVer41.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer41, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay41
			// 
			this.lblDay41.AutoSize = true;
			this.lblDay41.BackColor = System.Drawing.Color.Transparent;
			this.lblDay41.Location = new System.Drawing.Point(16, 16);
			this.lblDay41.Name = "lblDay41";
			this.lblDay41.Size = new System.Drawing.Size(0, 16);
			this.lblDay41.TabIndex = 61;
			this.lblDay41.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL40
			// 
			this.grpBL40.Controls.Add(this.txtNotas40);
			this.grpBL40.Controls.Add(this.btnVer40);
			this.grpBL40.Controls.Add(this.lblDay40);
			this.grpBL40.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL40.Location = new System.Drawing.Point(752, 592);
			this.grpBL40.Name = "grpBL40";
			this.grpBL40.Size = new System.Drawing.Size(184, 112);
			this.grpBL40.TabIndex = 46;
			this.grpBL40.TabStop = false;
			// 
			// txtNotas40
			// 
			this.txtNotas40.LinkColor = System.Drawing.Color.Black;
			this.txtNotas40.Location = new System.Drawing.Point(6, 15);
			this.txtNotas40.Name = "txtNotas40";
			this.txtNotas40.Size = new System.Drawing.Size(172, 82);
			this.txtNotas40.TabIndex = 767;
			// 
			// btnVer40
			// 
			this.btnVer40.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer40.Location = new System.Drawing.Point(8, 98);
			this.btnVer40.Name = "btnVer40";
			this.btnVer40.Size = new System.Drawing.Size(168, 8);
			this.btnVer40.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer40, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay40
			// 
			this.lblDay40.AutoSize = true;
			this.lblDay40.BackColor = System.Drawing.Color.Transparent;
			this.lblDay40.Location = new System.Drawing.Point(16, 16);
			this.lblDay40.Name = "lblDay40";
			this.lblDay40.Size = new System.Drawing.Size(0, 16);
			this.lblDay40.TabIndex = 61;
			this.lblDay40.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL39
			// 
			this.grpBL39.Controls.Add(this.txtNotas39);
			this.grpBL39.Controls.Add(this.btnVer39);
			this.grpBL39.Controls.Add(this.lblDay39);
			this.grpBL39.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL39.Location = new System.Drawing.Point(568, 592);
			this.grpBL39.Name = "grpBL39";
			this.grpBL39.Size = new System.Drawing.Size(184, 112);
			this.grpBL39.TabIndex = 45;
			this.grpBL39.TabStop = false;
			// 
			// txtNotas39
			// 
			this.txtNotas39.LinkColor = System.Drawing.Color.Black;
			this.txtNotas39.Location = new System.Drawing.Point(6, 15);
			this.txtNotas39.Name = "txtNotas39";
			this.txtNotas39.Size = new System.Drawing.Size(172, 82);
			this.txtNotas39.TabIndex = 767;
			// 
			// btnVer39
			// 
			this.btnVer39.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer39.Location = new System.Drawing.Point(8, 98);
			this.btnVer39.Name = "btnVer39";
			this.btnVer39.Size = new System.Drawing.Size(168, 8);
			this.btnVer39.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer39, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay39
			// 
			this.lblDay39.AutoSize = true;
			this.lblDay39.BackColor = System.Drawing.Color.Transparent;
			this.lblDay39.Location = new System.Drawing.Point(16, 16);
			this.lblDay39.Name = "lblDay39";
			this.lblDay39.Size = new System.Drawing.Size(0, 16);
			this.lblDay39.TabIndex = 61;
			this.lblDay39.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL38
			// 
			this.grpBL38.Controls.Add(this.txtNotas38);
			this.grpBL38.Controls.Add(this.btnVer38);
			this.grpBL38.Controls.Add(this.lblDay38);
			this.grpBL38.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL38.Location = new System.Drawing.Point(384, 592);
			this.grpBL38.Name = "grpBL38";
			this.grpBL38.Size = new System.Drawing.Size(184, 112);
			this.grpBL38.TabIndex = 44;
			this.grpBL38.TabStop = false;
			// 
			// txtNotas38
			// 
			this.txtNotas38.LinkColor = System.Drawing.Color.Black;
			this.txtNotas38.Location = new System.Drawing.Point(6, 15);
			this.txtNotas38.Name = "txtNotas38";
			this.txtNotas38.Size = new System.Drawing.Size(172, 82);
			this.txtNotas38.TabIndex = 767;
			// 
			// btnVer38
			// 
			this.btnVer38.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer38.Location = new System.Drawing.Point(8, 98);
			this.btnVer38.Name = "btnVer38";
			this.btnVer38.Size = new System.Drawing.Size(168, 8);
			this.btnVer38.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer38, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay38
			// 
			this.lblDay38.AutoSize = true;
			this.lblDay38.BackColor = System.Drawing.Color.Transparent;
			this.lblDay38.Location = new System.Drawing.Point(16, 16);
			this.lblDay38.Name = "lblDay38";
			this.lblDay38.Size = new System.Drawing.Size(0, 16);
			this.lblDay38.TabIndex = 61;
			this.lblDay38.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL37
			// 
			this.grpBL37.Controls.Add(this.txtNotas37);
			this.grpBL37.Controls.Add(this.btnVer37);
			this.grpBL37.Controls.Add(this.lblDay37);
			this.grpBL37.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL37.Location = new System.Drawing.Point(200, 592);
			this.grpBL37.Name = "grpBL37";
			this.grpBL37.Size = new System.Drawing.Size(184, 112);
			this.grpBL37.TabIndex = 43;
			this.grpBL37.TabStop = false;
			// 
			// txtNotas37
			// 
			this.txtNotas37.LinkColor = System.Drawing.Color.Black;
			this.txtNotas37.Location = new System.Drawing.Point(6, 15);
			this.txtNotas37.Name = "txtNotas37";
			this.txtNotas37.Size = new System.Drawing.Size(172, 82);
			this.txtNotas37.TabIndex = 766;
			// 
			// btnVer37
			// 
			this.btnVer37.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer37.Location = new System.Drawing.Point(8, 98);
			this.btnVer37.Name = "btnVer37";
			this.btnVer37.Size = new System.Drawing.Size(168, 8);
			this.btnVer37.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer37, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay37
			// 
			this.lblDay37.AutoSize = true;
			this.lblDay37.BackColor = System.Drawing.Color.Transparent;
			this.lblDay37.Location = new System.Drawing.Point(16, 16);
			this.lblDay37.Name = "lblDay37";
			this.lblDay37.Size = new System.Drawing.Size(0, 16);
			this.lblDay37.TabIndex = 61;
			this.lblDay37.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL36
			// 
			this.grpBL36.Controls.Add(this.txtNotas36);
			this.grpBL36.Controls.Add(this.btnVer36);
			this.grpBL36.Controls.Add(this.lblDay36);
			this.grpBL36.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL36.Location = new System.Drawing.Point(16, 592);
			this.grpBL36.Name = "grpBL36";
			this.grpBL36.Size = new System.Drawing.Size(184, 112);
			this.grpBL36.TabIndex = 42;
			this.grpBL36.TabStop = false;
			this.grpBL36.Enter += new System.EventHandler(this.groupBox49_Enter);
			// 
			// txtNotas36
			// 
			this.txtNotas36.LinkColor = System.Drawing.Color.Black;
			this.txtNotas36.Location = new System.Drawing.Point(6, 15);
			this.txtNotas36.Name = "txtNotas36";
			this.txtNotas36.Size = new System.Drawing.Size(172, 82);
			this.txtNotas36.TabIndex = 767;
			// 
			// btnVer36
			// 
			this.btnVer36.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer36.Location = new System.Drawing.Point(8, 98);
			this.btnVer36.Name = "btnVer36";
			this.btnVer36.Size = new System.Drawing.Size(168, 8);
			this.btnVer36.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer36, "Click Para Visualizar Proyectos del Día");
			// 
			// lblDay36
			// 
			this.lblDay36.AutoSize = true;
			this.lblDay36.BackColor = System.Drawing.Color.Transparent;
			this.lblDay36.Location = new System.Drawing.Point(16, 16);
			this.lblDay36.Name = "lblDay36";
			this.lblDay36.Size = new System.Drawing.Size(0, 16);
			this.lblDay36.TabIndex = 61;
			this.lblDay36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(16, 12);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(24, 16);
			this.label8.TabIndex = 752;
			this.label8.Text = "Año";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtAño
			// 
			this.txtAño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAño.Location = new System.Drawing.Point(48, 10);
			this.txtAño.MinValue = 2014;
			this.txtAño.Name = "txtAño";
			this.txtAño.PromptChar = ' ';
			this.txtAño.Size = new System.Drawing.Size(88, 21);
			this.txtAño.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAño.TabIndex = 751;
			this.txtAño.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAño_KeyDown);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(144, 12);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(26, 16);
			this.label9.TabIndex = 754;
			this.label9.Text = "Mes";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMes
			// 
			this.txtMes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMes.Location = new System.Drawing.Point(184, 10);
			this.txtMes.MaxValue = 12;
			this.txtMes.MinValue = 1;
			this.txtMes.Name = "txtMes";
			this.txtMes.PromptChar = ' ';
			this.txtMes.Size = new System.Drawing.Size(88, 21);
			this.txtMes.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtMes.TabIndex = 755;
			this.txtMes.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtMes_KeyDown);
			this.txtMes.ValueChanged += new System.EventHandler(this.txtMes_ValueChanged);
			// 
			// lblMes
			// 
			this.lblMes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblMes.AutoSize = true;
			this.lblMes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblMes.ForeColor = System.Drawing.Color.Firebrick;
			this.lblMes.Location = new System.Drawing.Point(504, -36);
			this.lblMes.Name = "lblMes";
			this.lblMes.Size = new System.Drawing.Size(0, 16);
			this.lblMes.TabIndex = 757;
			this.lblMes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(456, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 758;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
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
			// lblMeses
			// 
			this.lblMeses.AutoSize = true;
			this.lblMeses.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblMeses.ForeColor = System.Drawing.Color.Firebrick;
			this.lblMeses.Location = new System.Drawing.Point(288, 8);
			this.lblMeses.Name = "lblMeses";
			this.lblMeses.Size = new System.Drawing.Size(0, 16);
			this.lblMeses.TabIndex = 759;
			this.lblMeses.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDay30
			// 
			this.lblDay30.AutoSize = true;
			this.lblDay30.BackColor = System.Drawing.Color.Transparent;
			this.lblDay30.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDay30.Location = new System.Drawing.Point(16, 16);
			this.lblDay30.Name = "lblDay30";
			this.lblDay30.Size = new System.Drawing.Size(0, 16);
			this.lblDay30.TabIndex = 61;
			this.lblDay30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// grpBL30
			// 
			this.grpBL30.Controls.Add(this.txtNotas30);
			this.grpBL30.Controls.Add(this.btnVer30);
			this.grpBL30.Controls.Add(this.lblDay30);
			this.grpBL30.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpBL30.Location = new System.Drawing.Point(200, 488);
			this.grpBL30.Name = "grpBL30";
			this.grpBL30.Size = new System.Drawing.Size(184, 104);
			this.grpBL30.TabIndex = 36;
			this.grpBL30.TabStop = false;
			// 
			// txtNotas30
			// 
			this.txtNotas30.LinkColor = System.Drawing.Color.Black;
			this.txtNotas30.Location = new System.Drawing.Point(6, 11);
			this.txtNotas30.Name = "txtNotas30";
			this.txtNotas30.Size = new System.Drawing.Size(172, 80);
			this.txtNotas30.TabIndex = 766;
			// 
			// btnVer30
			// 
			this.btnVer30.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer30.Location = new System.Drawing.Point(8, 92);
			this.btnVer30.Name = "btnVer30";
			this.btnVer30.Size = new System.Drawing.Size(168, 8);
			this.btnVer30.TabIndex = 763;
			this.toolTip1.SetToolTip(this.btnVer30, "Click Para Visualizar Proyectos del Día");
			// 
			// button1
			// 
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(568, 8);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(75, 24);
			this.button1.TabIndex = 760;
			this.button1.Text = "Ver";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// frmProyectosZoomatica
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1304, 710);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.lblMeses);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.lblMes);
			this.Controls.Add(this.txtMes);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtAño);
			this.Controls.Add(this.grpBL42);
			this.Controls.Add(this.grpBL41);
			this.Controls.Add(this.grpBL40);
			this.Controls.Add(this.grpBL39);
			this.Controls.Add(this.grpBL38);
			this.Controls.Add(this.grpBL37);
			this.Controls.Add(this.grpBL36);
			this.Controls.Add(this.grpBL35);
			this.Controls.Add(this.grpBL34);
			this.Controls.Add(this.grpBL33);
			this.Controls.Add(this.grpBL32);
			this.Controls.Add(this.grpBL31);
			this.Controls.Add(this.grpBL30);
			this.Controls.Add(this.grpBL29);
			this.Controls.Add(this.grpBL28);
			this.Controls.Add(this.grpBL21);
			this.Controls.Add(this.grpBL14);
			this.Controls.Add(this.grpBL7);
			this.Controls.Add(this.grpBL27);
			this.Controls.Add(this.grpBL20);
			this.Controls.Add(this.grpBL9);
			this.Controls.Add(this.grpBL16);
			this.Controls.Add(this.grpBL23);
			this.Controls.Add(this.grpBL3);
			this.Controls.Add(this.grpBL24);
			this.Controls.Add(this.grpBL17);
			this.Controls.Add(this.grpBL10);
			this.Controls.Add(this.grpBL4);
			this.Controls.Add(this.grpBL11);
			this.Controls.Add(this.grpBL18);
			this.Controls.Add(this.grpBL25);
			this.Controls.Add(this.grpBL5);
			this.Controls.Add(this.grpBL12);
			this.Controls.Add(this.grpBL19);
			this.Controls.Add(this.grpBL26);
			this.Controls.Add(this.grpBL13);
			this.Controls.Add(this.grpBL6);
			this.Controls.Add(this.grpBL2);
			this.Controls.Add(this.grpBL22);
			this.Controls.Add(this.grpBL15);
			this.Controls.Add(this.grpBL8);
			this.Controls.Add(this.grpBL1);
			this.Controls.Add(this.groupBox7);
			this.Controls.Add(this.groupBox6);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmProyectosZoomatica";
			this.Text = "Cronograma De Proyectos";
			this.Load += new System.EventHandler(this.frmProyectosZoomatica_Load);
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			this.groupBox5.ResumeLayout(false);
			this.groupBox6.ResumeLayout(false);
			this.groupBox7.ResumeLayout(false);
			this.grpBL1.ResumeLayout(false);
			this.grpBL8.ResumeLayout(false);
			this.grpBL15.ResumeLayout(false);
			this.grpBL22.ResumeLayout(false);
			this.grpBL2.ResumeLayout(false);
			this.grpBL6.ResumeLayout(false);
			this.grpBL13.ResumeLayout(false);
			this.grpBL26.ResumeLayout(false);
			this.grpBL19.ResumeLayout(false);
			this.grpBL12.ResumeLayout(false);
			this.grpBL5.ResumeLayout(false);
			this.grpBL25.ResumeLayout(false);
			this.grpBL18.ResumeLayout(false);
			this.grpBL11.ResumeLayout(false);
			this.grpBL4.ResumeLayout(false);
			this.grpBL10.ResumeLayout(false);
			this.grpBL17.ResumeLayout(false);
			this.grpBL24.ResumeLayout(false);
			this.grpBL3.ResumeLayout(false);
			this.grpBL23.ResumeLayout(false);
			this.grpBL16.ResumeLayout(false);
			this.grpBL9.ResumeLayout(false);
			this.grpBL27.ResumeLayout(false);
			this.grpBL20.ResumeLayout(false);
			this.grpBL28.ResumeLayout(false);
			this.grpBL21.ResumeLayout(false);
			this.grpBL14.ResumeLayout(false);
			this.grpBL7.ResumeLayout(false);
			this.grpBL29.ResumeLayout(false);
			this.grpBL31.ResumeLayout(false);
			this.grpBL32.ResumeLayout(false);
			this.grpBL33.ResumeLayout(false);
			this.grpBL34.ResumeLayout(false);
			this.grpBL35.ResumeLayout(false);
			this.grpBL42.ResumeLayout(false);
			this.grpBL41.ResumeLayout(false);
			this.grpBL40.ResumeLayout(false);
			this.grpBL39.ResumeLayout(false);
			this.grpBL38.ResumeLayout(false);
			this.grpBL37.ResumeLayout(false);
			this.grpBL36.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.grpBL30.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void label28_Click(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox49_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void frmProyectosZoomatica_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0630");

			if (!Funcion.Permiso("1219", cdsSeteoF))
			{				
				MessageBox.Show("No Tiene Acceso a Calendario de Proyectos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.txtAño.Value = DateTime.Today.Year;
			this.txtMes.Value = DateTime.Today.Month;
			this.txtMes_ValueChanged (sender, e);
		}

		private void txtAño_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				this.txtAño.SelectAll ();
		}

		private void txtMes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			this.txtMes.SelectAll ();
		}

		private void txtMes_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int) this.txtMes.Value == 1) this.lblMeses.Text = "ENERO";
			if ((int) this.txtMes.Value == 2) this.lblMeses.Text = "FEBRERO";
			if ((int) this.txtMes.Value == 3) this.lblMeses.Text = "MARZO";
			if ((int) this.txtMes.Value == 4) this.lblMeses.Text = "ABRIL";
			if ((int) this.txtMes.Value == 5) this.lblMeses.Text = "MAYO";
			if ((int) this.txtMes.Value == 6) this.lblMeses.Text = "JUNIO";
			if ((int) this.txtMes.Value == 7) this.lblMeses.Text = "JULIO";
			if ((int) this.txtMes.Value == 8) this.lblMeses.Text = "AGOSTO";
			if ((int) this.txtMes.Value == 9) this.lblMeses.Text = "SEPTIEMBRE";
			if ((int) this.txtMes.Value == 10) this.lblMeses.Text = "OCTUBRE";
			if ((int) this.txtMes.Value == 11) this.lblMeses.Text = "NOVIEMBRE";
			if ((int) this.txtMes.Value == 12) this.lblMeses.Text = "DICIEMBRE";

			this.btnVer_Click (sender, e);

		}
     int iBotones =0;

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Vaciar ();

			int iDiaA = Convert.ToInt32( ( Convert.ToDateTime(DateTime.Today  ).Day));
			int iMesA = (int)this.txtMes.Value;
			int iYearA = (int)this.txtAño.Value;
			int iMesB = Convert.ToInt32( ( Convert.ToDateTime(DateTime.Today ).Month));
			int iYearB = Convert.ToInt32( ( Convert.ToDateTime(DateTime.Today).Year));

//			Button btn1 = new Button(); 
//			btn1.Location=new System.Drawing.Point(8,16); 
//			btn1.Size = new System.Drawing.Size(168,24); 
//			
//			btn1.Name = "btn2";
//			btn1.Text = btn1.Name;
//		
//			this.Controls.Add(btn1);
//			this.grpBL1.Controls.Add(btn1);

			bool DiaActual = false;
    
			if ( iMesA == iMesB && iYearA == iYearB)
			{
         DiaActual = true;
			}
			int iDiaSemana = 0;
			int iUltimodia = 0;
			int iContActividades = 0;

			DateTime dtInicio = DateTime.Parse("01/" + (int)this.txtMes.Value+ "/" + (int)this.txtAño.Value);
			iDiaSemana = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" select DATEPART(dw, '{0}')", Convert.ToDateTime(dtInicio).ToString("yyyyMMdd")));		
      iUltimodia = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" SELECT DATEPART( DAY,DATEADD(s,-1,DATEADD(mm, DATEDIFF(m,0,'{0}')+1,0)))", Convert.ToDateTime(dtInicio).ToString("yyyyMMdd")));	

			if (iDiaSemana == 2) this.grpBL1.Text = "1";
			if (iDiaSemana == 3) this.grpBL2.Text = "1";
			if (iDiaSemana == 4) this.grpBL3.Text = "1";
			if (iDiaSemana == 5) this.grpBL4.Text = "1";
			if (iDiaSemana == 6) this.grpBL5.Text = "1";
			if (iDiaSemana == 7) this.grpBL6.Text = "1";
			if (iDiaSemana == 1) this.grpBL7.Text = "1";
    
			int iCont = 0;
			int iDiaMes = 0;
			int i =0;
			int j =0;
			int iPosX = 8;
			int iPosY = 13;
			
			string sTarea ="";

			GroupBox[] aGroupBox = { grpBL1, grpBL2, grpBL3, grpBL4, grpBL5, grpBL6, grpBL7, grpBL8,grpBL9, grpBL10
												 , grpBL11, grpBL12,grpBL13, grpBL14, grpBL15, grpBL16,grpBL17, grpBL18, grpBL19, grpBL20,
												 grpBL21, grpBL22, grpBL23, grpBL24,grpBL25, grpBL26, grpBL27, grpBL28,grpBL29, grpBL30, grpBL31, grpBL32,
												 grpBL33, grpBL34, grpBL35, grpBL36, grpBL37, grpBL38, grpBL39, grpBL40, grpBL41, grpBL42 };

			Button[] aButton = { btnVer1, btnVer2, btnVer3, btnVer4, btnVer5, btnVer6, btnVer7, btnVer8, btnVer9, btnVer10
															 , btnVer11, btnVer12, btnVer13, btnVer14, btnVer15, btnVer16,btnVer17, btnVer18, btnVer19, btnVer20,
															 btnVer21, btnVer22, btnVer23, btnVer24,btnVer25, btnVer26, btnVer27, btnVer28, btnVer29, btnVer30, btnVer31, btnVer32,
															 btnVer33, btnVer34, btnVer35, btnVer36, btnVer37, btnVer38, btnVer39, btnVer40, btnVer41, btnVer42 };
 
	LinkLabel [] aText= {  txtNotas1, txtNotas2, txtNotas3, txtNotas4, txtNotas5, txtNotas6, txtNotas7, txtNotas8, txtNotas9, txtNotas10
															 , txtNotas11, txtNotas12,txtNotas13, txtNotas14, txtNotas15, txtNotas16,txtNotas17, txtNotas18, txtNotas19, txtNotas20,
															 txtNotas21, txtNotas22, txtNotas23, txtNotas24,txtNotas25, txtNotas26, txtNotas27, txtNotas28, txtNotas29, txtNotas30, txtNotas31, txtNotas32,
															 txtNotas33, txtNotas34, txtNotas35, txtNotas36, txtNotas37, txtNotas38, txtNotas39, txtNotas40, txtNotas41, txtNotas42 };
			
			for ( i = 0; i < aGroupBox.Length; i++)
			{
				if (aGroupBox[i].Text =="1")
				{
					j=i;
					for ( iDiaMes = 1; iDiaMes <= iUltimodia; iDiaMes++)
					{
						aGroupBox[j].Text = ""+ iDiaMes;
						DateTime dtFecha  = DateTime.Parse(iDiaMes + "/" + iMesA + "/" + iYearA);
						sTarea = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" DECLARE @valores VARCHAR(1000) ='' SELECT @valores= COALESCE(@valores + '*- ', '') + descripcion FROM facturaPromociones where convert (date, Desde) = convert (date, '{0}') select @valores as valores ", Convert.ToDateTime(dtFecha).ToString("yyyyMMdd")));	
						string sTamano = sTarea ;
						iContActividades = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" declare @Count int = 0; select @Count = COUNT (*) from FacturaPromociones where Convert (Date, '{0}') between Convert (Date, Desde) and Convert (Date, Hasta) Select @Count", Convert.ToDateTime(dtFecha).ToString("yyyyMMdd")));	

						if ( sTamano.Length > 1 )
						{
							aText[j].Visible = true;
							aButton[j].Visible = true;
							aText[j].Text = sTarea;
						}
						else
						{
							aText[j].Visible = false;
							aButton[j].Visible = false;
						}
						if (iContActividades > 0)
						{
							for ( int R  = 1; R <= iContActividades; R ++)
							{
								iBotones = iBotones + 1;
								string sNomBo ="btn"+iBotones;
								Button btn1 = new Button(); 
								btn1.Location=new System.Drawing.Point(iPosX,iPosY); 
								btn1.Size = new System.Drawing.Size(168,16); 
//								btn1.Name = aGroupBox[j].Text;
								btn1.Text = ""+R;
                int idcontdr = FuncionesProcedimientos.dtGeneral(string.Format("SELECT Descripcion AS valores  FROM facturaPromociones where convert (date, Desde) = convert (date, '{0}')", Convert.ToDateTime(dtFecha).ToString("yyyyMMdd"))).Rows.Count;
								foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec ConsultaActividadesZoomatika '{0}'", Convert.ToDateTime(dtFecha).ToString("yyyyMMdd"))).Rows)
								{
									try
									{
										if (Convert.ToInt32 ( dr["IdPlanificacionPagos"].ToString()) ==  Convert.ToInt32(R))
										{
										
											btn1.Text = dr["valores"].ToString();
											toolTip1.SetToolTip(btn1,dr["valores"].ToString());
										}
						
									}						
																		catch (Exception exc)
																		{
										MessageBox.Show(exc.Message + " erorr : " );
																			continue;
																		}
								}
								
							
								toolTip2.SetToolTip(btn1,aGroupBox[j].Text);
								toolTip3.SetToolTip(btn1,aGroupBox[j].Name);
								
							
								
								this.Controls.Add(btn1);
								aGroupBox[j].Controls.Add(btn1);
															   btn1.Click += new EventHandler(btn_Click);  
										
								iPosY = iPosY + 18;
							}
							iPosX = 8;
							iPosY = 13;
							iContActividades = 0;
						}

						#region Pinta dia actual 
						if (DiaActual) 
						{
							if (iDiaMes == iDiaA )
							{
								aGroupBox[j].BackColor = Color.Green ;
								aText[j].Visible = true ;
								aText[j].BackColor = Color.Green;
								aButton[j].BackColor = Color.White;
						

							}
							else
							{
								aGroupBox[j].BackColor = Color.Transparent ;
							}
						}
						#endregion Pinta dia actual 

						j= j+1;
					}
				}
			}

			#region Oculta dias que no tienen fecha 

			for ( int a = 0; a < aGroupBox.Length; a++)
			{
				if (aGroupBox[a].Text.Length == 0 )
				{
					aGroupBox[a].Visible = false ;
				}
				else
				{
					aGroupBox[a].Visible = true ;
				}
			}
			#endregion Oculta dias que no tienen fecha 


			if ((int) this.txtMes.Value == 1) this.lblMes.Text = "ENERO";
			if ((int) this.txtMes.Value == 2) this.lblMes.Text = "FEBRERO";
			if ((int) this.txtMes.Value == 3) this.lblMes.Text = "MARZO";
			if ((int) this.txtMes.Value == 4) this.lblMes.Text = "ABRIL";
			if ((int) this.txtMes.Value == 5) this.lblMes.Text = "MAYO";
			if ((int) this.txtMes.Value == 6) this.lblMes.Text = "JUNIO";
			if ((int) this.txtMes.Value == 7) this.lblMes.Text = "JULIO";
			if ((int) this.txtMes.Value == 8) this.lblMes.Text = "AGOSTO";
			if ((int) this.txtMes.Value == 9) this.lblMes.Text = "SEPTIEMBRE";
			if ((int) this.txtMes.Value == 10) this.lblMes.Text = "OCTUBRE";
			if ((int) this.txtMes.Value == 11) this.lblMes.Text = "NOVIEMBRE";
			if ((int) this.txtMes.Value == 12) this.lblMes.Text = "DICIEMBRE";
//			this.btn_Click (sender, e);
			
		}

		private void btn_Click(object sender, EventArgs e)
		{
			
			GroupBox[] aGroupBox = { grpBL1, grpBL2, grpBL3, grpBL4, grpBL5, grpBL6, grpBL7, grpBL8,grpBL9, grpBL10
															 , grpBL11, grpBL12,grpBL13, grpBL14, grpBL15, grpBL16,grpBL17, grpBL18, grpBL19, grpBL20,
															 grpBL21, grpBL22, grpBL23, grpBL24,grpBL25, grpBL26, grpBL27, grpBL28,grpBL29, grpBL30, grpBL31, grpBL32,
															 grpBL33, grpBL34, grpBL35, grpBL36, grpBL37, grpBL38, grpBL39, grpBL40, grpBL41, grpBL42 };
			
			
			string texto = ((Button)sender).Name;
			string texto2 = ((Button)sender).Text;
			for (int i = 0; i < aGroupBox.Length; i++)
			{
				if ( Convert.ToString( texto) == Convert.ToString( aGroupBox[i].Text) )
				{
					int iCont =0;
					foreach(Control c in aGroupBox[Convert.ToInt32(i) ].Controls)
					{
						
						if(c is Button)
						{
							c.BackColor = Color.Red;
							c.Name = c.Name;
							iCont = iCont +1 ;
							c.Text = ""+iCont;
						}
					}
					
				}
			}
			
		}

	
		public void Vaciar ()
		{
			GroupBox[] aGroupBox = { grpBL1, grpBL2, grpBL3, grpBL4, grpBL5, grpBL6, grpBL7, grpBL8,grpBL9, grpBL10
															 , grpBL11, grpBL12,grpBL13, grpBL14, grpBL15, grpBL16,grpBL17, grpBL18, grpBL19, grpBL20,
															 grpBL21, grpBL22, grpBL23, grpBL24,grpBL25, grpBL26, grpBL27, grpBL28,grpBL29, grpBL30, grpBL31, grpBL32,
															 grpBL33, grpBL34, grpBL35, grpBL36, grpBL37, grpBL38, grpBL39, grpBL40, grpBL41, grpBL42 };
			
			LinkLabel [] aText= { txtNotas1, txtNotas2, txtNotas3, txtNotas4, txtNotas5, txtNotas6, txtNotas7, txtNotas8, txtNotas9, txtNotas10
																																		, txtNotas11, txtNotas12,txtNotas13, txtNotas14, txtNotas15, txtNotas16,txtNotas17, txtNotas18, txtNotas19, txtNotas20,
																																		txtNotas21, txtNotas22, txtNotas23, txtNotas24,txtNotas25, txtNotas26, txtNotas27, txtNotas28, txtNotas29, txtNotas30, txtNotas31, txtNotas32,
																																		txtNotas33, txtNotas34, txtNotas35, txtNotas36, txtNotas37, txtNotas38, txtNotas39, txtNotas40, txtNotas41, txtNotas42 };
			
			
			for ( int i = 0; i < aGroupBox.Length; i++)
			{
				  aGroupBox[i].Text = "";
				  aGroupBox[i].BackColor = Color.Transparent ;
				  aText[i].Text ="";
			    aGroupBox[i].Controls.Clear();
			}

		}
		public void evento()
		{
	
			MessageBox.Show("");
		}


		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void ultraTextEditor39_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void linkLabel12_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
		
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
				MessageBox.Show("Prueba de sistemas", "Información", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//			string boton =this.button1.Text;
		}
	}
}
