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
	/// Descripción breve de frmPromocionesLote.
	/// </summary>
	public class frmPromocionesLote : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.TextBox txtNomPromocion;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ChkDias;
		private Infragistics.Win.UltraWinGrid.UltraGrid uDiasSemana;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ChkHabitaciones;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridHoras;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugHabitaciones;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label5;
		private C1.Data.C1DataSet cdsBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFeIni;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFeFin;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.RadioButton rbFijo;
		private System.Windows.Forms.RadioButton rbDescuento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.CheckBox checkBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.CheckBox checkBox2;
		public Infragistics.Win.UltraWinGrid.UltraGrid ugArtPromo;
		private System.Windows.Forms.Button button1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPromocionesLote()
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

		#region VariablesSis
		bool activo = false;
		bool actv = false;
		int laBodega = 0;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ChkHoras;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCant;
		int idPromo = 0;
		int numHabtc = 0;
    int numDia = 0;
		int idArtPromo = 0;
		#endregion VariablesSis

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPromocionesLote));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDiaPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacionPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activa");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDiaPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacionPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activa");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHoraPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDiaPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHora");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorTiempo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activa");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHoraPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDiaPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorTiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activa");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacionPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tematica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroHabitacion");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activa");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacionPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tematica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activa");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromoArt");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtNomPromocion = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.dtFeIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFeFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.ChkDias = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uDiasSemana = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ChkHabitaciones = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridHoras = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ugHabitaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ChkHoras = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.cdsBodega = new C1.Data.C1DataSet();
			this.txtCant = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.rbDescuento = new System.Windows.Forms.RadioButton();
			this.rbFijo = new System.Windows.Forms.RadioButton();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.checkBox1 = new System.Windows.Forms.CheckBox();
			this.ugArtPromo = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.checkBox2 = new System.Windows.Forms.CheckBox();
			this.button1 = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uDiasSemana)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridHoras)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCant)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugArtPromo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.SuspendLayout();
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.Location = new System.Drawing.Point(16, 5);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(122, 16);
			this.lblBodega.TabIndex = 132;
			this.lblBodega.Text = "Nombre PROMOCION";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNomPromocion
			// 
			this.txtNomPromocion.AcceptsReturn = true;
			this.txtNomPromocion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.txtNomPromocion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNomPromocion.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNomPromocion.Location = new System.Drawing.Point(12, 21);
			this.txtNomPromocion.Name = "txtNomPromocion";
			this.txtNomPromocion.Size = new System.Drawing.Size(296, 26);
			this.txtNomPromocion.TabIndex = 232;
			this.txtNomPromocion.Text = "";
			this.txtNomPromocion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNomPromocion_KeyDown);
			this.txtNomPromocion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNomPromocion_KeyPress);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(504, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(65, 16);
			this.label1.TabIndex = 233;
			this.label1.Text = "Fecha Inicia";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(648, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(78, 16);
			this.label2.TabIndex = 235;
			this.label2.Text = "Fecha Finaliza";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.AcceptsReturn = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance1;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(12, 68);
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(444, 36);
			this.txtDescripcion.TabIndex = 237;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(12, 52);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(121, 16);
			this.label3.TabIndex = 238;
			this.label3.Text = "Descripción Promoción";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFeIni
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFeIni.Appearance = appearance2;
			this.dtFeIni.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.dtFeIni.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFeIni.DateButtons.Add(dateButton1);
			this.dtFeIni.Format = "dd/MM/yyyy HH:mm";
			this.dtFeIni.Location = new System.Drawing.Point(504, 24);
			this.dtFeIni.Name = "dtFeIni";
			this.dtFeIni.NonAutoSizeHeight = 23;
			this.dtFeIni.Size = new System.Drawing.Size(142, 21);
			this.dtFeIni.SpinButtonsVisible = true;
			this.dtFeIni.TabIndex = 239;
			this.dtFeIni.Value = ((object)(resources.GetObject("dtFeIni.Value")));
			// 
			// dtFeFin
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFeFin.Appearance = appearance3;
			this.dtFeFin.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.dtFeFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFeFin.DateButtons.Add(dateButton2);
			this.dtFeFin.Format = "dd/MM/yyyy HH:mm";
			this.dtFeFin.Location = new System.Drawing.Point(648, 24);
			this.dtFeFin.Name = "dtFeFin";
			this.dtFeFin.NonAutoSizeHeight = 23;
			this.dtFeFin.Size = new System.Drawing.Size(142, 21);
			this.dtFeFin.SpinButtonsVisible = true;
			this.dtFeFin.TabIndex = 240;
			this.dtFeFin.Value = ((object)(resources.GetObject("dtFeFin.Value")));
			this.dtFeFin.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFeFin_KeyPress);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(752, 106);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 16);
			this.label4.TabIndex = 241;
			this.label4.Text = "# Horas";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Location = new System.Drawing.Point(16, 544);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(136, 32);
			this.btnGuardar.TabIndex = 243;
			this.btnGuardar.Text = " GUARDAR";
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// ChkDias
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ChkDias.Appearance = appearance4;
			this.ChkDias.BackColor = System.Drawing.Color.Transparent;
			this.ChkDias.CausesValidation = false;
			this.ChkDias.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.ChkDias.Location = new System.Drawing.Point(664, 240);
			this.ChkDias.Name = "ChkDias";
			this.ChkDias.Size = new System.Drawing.Size(16, 8);
			this.ChkDias.TabIndex = 321;
			this.ChkDias.CheckedChanged += new System.EventHandler(this.ChkDias_CheckedChanged);
			// 
			// uDiasSemana
			// 
			this.uDiasSemana.Cursor = System.Windows.Forms.Cursors.Default;
			this.uDiasSemana.DataSource = this.ultraDataSource2;
			appearance5.BackColor = System.Drawing.Color.White;
			this.uDiasSemana.DisplayLayout.Appearance = appearance5;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 46;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 220;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn5.Width = 30;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			ultraGridBand1.Header.Caption = "DIAS SEMANA";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance6.FontData.SizeInPoints = 7.25F;
			ultraGridBand1.Override.HeaderAppearance = appearance6;
			this.uDiasSemana.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uDiasSemana.DisplayLayout.Override.ActiveRowAppearance = appearance7;
			this.uDiasSemana.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uDiasSemana.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.uDiasSemana.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 10F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uDiasSemana.DisplayLayout.Override.HeaderAppearance = appearance9;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uDiasSemana.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uDiasSemana.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uDiasSemana.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.uDiasSemana.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uDiasSemana.Location = new System.Drawing.Point(420, 232);
			this.uDiasSemana.Name = "uDiasSemana";
			this.uDiasSemana.Size = new System.Drawing.Size(280, 304);
			this.uDiasSemana.TabIndex = 320;
			this.uDiasSemana.Click += new System.EventHandler(this.uDiasSemana_Click);
			// 
			// ultraDataSource2
			// 
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// ChkHabitaciones
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.ChkHabitaciones.Appearance = appearance13;
			this.ChkHabitaciones.BackColor = System.Drawing.Color.Transparent;
			this.ChkHabitaciones.CausesValidation = false;
			this.ChkHabitaciones.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.ChkHabitaciones.Location = new System.Drawing.Point(368, 240);
			this.ChkHabitaciones.Name = "ChkHabitaciones";
			this.ChkHabitaciones.Size = new System.Drawing.Size(16, 8);
			this.ChkHabitaciones.TabIndex = 319;
			this.ChkHabitaciones.CheckedChanged += new System.EventHandler(this.ChkHabitaciones_CheckedChanged);
			// 
			// uGridHoras
			// 
			this.uGridHoras.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridHoras.DataSource = this.ultraDataSource3;
			appearance14.BackColor = System.Drawing.Color.White;
			this.uGridHoras.DisplayLayout.Appearance = appearance14;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Width = 190;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance15;
			ultraGridColumn10.Format = "#,##0";
			ultraGridColumn10.Header.Caption = "Valor";
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 30;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn11.Width = 30;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			ultraGridBand2.Header.Caption = "HORAS";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance16.FontData.SizeInPoints = 7.25F;
			ultraGridBand2.Override.HeaderAppearance = appearance16;
			this.uGridHoras.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHoras.DisplayLayout.Override.ActiveRowAppearance = appearance17;
			this.uGridHoras.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridHoras.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.uGridHoras.DisplayLayout.Override.CardAreaAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 10F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridHoras.DisplayLayout.Override.HeaderAppearance = appearance19;
			appearance20.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance20.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHoras.DisplayLayout.Override.RowAlternateAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHoras.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHoras.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.uGridHoras.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridHoras.Location = new System.Drawing.Point(704, 232);
			this.uGridHoras.Name = "uGridHoras";
			this.uGridHoras.Size = new System.Drawing.Size(136, 304);
			this.uGridHoras.TabIndex = 318;
			this.uGridHoras.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(bool);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// ugHabitaciones
			// 
			this.ugHabitaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugHabitaciones.DataSource = this.ultraDataSource1;
			appearance23.BackColor = System.Drawing.Color.White;
			this.ugHabitaciones.DisplayLayout.Appearance = appearance23;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.Caption = "Habitación";
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 341;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn15.Header.Caption = "Nombre";
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Width = 170;
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Width = 110;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellAppearance = appearance24;
			ultraGridColumn17.Header.Caption = "Num";
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn17.Width = 50;
			ultraGridColumn18.Header.Caption = "Sel";
			ultraGridColumn18.Header.VisiblePosition = 6;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			ultraGridBand3.Header.Caption = "HABITACIONES";
			ultraGridBand3.HeaderVisible = true;
			this.ugHabitaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugHabitaciones.DisplayLayout.Override.ActiveRowAppearance = appearance25;
			this.ugHabitaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugHabitaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugHabitaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance26.BackColor = System.Drawing.Color.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.CardAreaAppearance = appearance26;
			this.ugHabitaciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance27.FontData.BoldAsString = "True";
			appearance27.FontData.Name = "Arial";
			appearance27.FontData.SizeInPoints = 10F;
			appearance27.ForeColor = System.Drawing.Color.White;
			appearance27.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.HeaderAppearance = appearance27;
			appearance28.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance28.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowAlternateAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.ugHabitaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugHabitaciones.Location = new System.Drawing.Point(12, 232);
			this.ugHabitaciones.Name = "ugHabitaciones";
			this.ugHabitaciones.Size = new System.Drawing.Size(392, 304);
			this.ugHabitaciones.TabIndex = 317;
			this.ugHabitaciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugHabitaciones_InitializeLayout);
			this.ugHabitaciones.Click += new System.EventHandler(this.ugHabitaciones_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(bool);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// ChkHoras
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.ChkHoras.Appearance = appearance31;
			this.ChkHoras.BackColor = System.Drawing.Color.Transparent;
			this.ChkHoras.CausesValidation = false;
			this.ChkHoras.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.ChkHoras.Location = new System.Drawing.Point(760, 240);
			this.ChkHoras.Name = "ChkHoras";
			this.ChkHoras.Size = new System.Drawing.Size(16, 8);
			this.ChkHoras.TabIndex = 322;
			this.ChkHoras.Visible = false;
			this.ChkHoras.CheckedChanged += new System.EventHandler(this.ChkHoras_CheckedChanged);
			// 
			// btBuscar
			// 
			this.btBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btBuscar.Image")));
			this.btBuscar.Location = new System.Drawing.Point(472, 23);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(24, 22);
			this.btBuscar.TabIndex = 325;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Width = 30;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 30;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Width = 30;
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(320, 24);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(152, 21);
			this.cmbBodega.TabIndex = 324;
			this.cmbBodega.ValueMember = "idBodega";
			this.cmbBodega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodega_KeyPress);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(320, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(39, 19);
			this.label5.TabIndex = 323;
			this.label5.Text = "Local";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cdsBodega
			// 
			this.cdsBodega.BindingContextCtrl = this;
			this.cdsBodega.DataLibrary = "LibFacturacion";
			this.cdsBodega.DataLibraryUrl = "";
			this.cdsBodega.DataSetDef = "dsBodega";
			this.cdsBodega.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsBodega.Name = "cdsBodega";
			this.cdsBodega.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsBodega.SchemaDef = null;
			this.cdsBodega.BeforeFill += new C1.Data.FillEventHandler(this.cdsBodega_BeforeFill);
			// 
			// txtCant
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCant.Appearance = appearance32;
			this.txtCant.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCant.FormatString = "#,##0";
			this.txtCant.Location = new System.Drawing.Point(800, 104);
			this.txtCant.Name = "txtCant";
			this.txtCant.PromptChar = ' ';
			this.txtCant.Size = new System.Drawing.Size(40, 21);
			this.txtCant.TabIndex = 328;
			this.txtCant.Click += new System.EventHandler(this.txtCant_Click);
			this.txtCant.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCant_KeyPress);
			this.txtCant.Enter += new System.EventHandler(this.txtCant_Enter);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.rbDescuento);
			this.groupBox1.Controls.Add(this.rbFijo);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(624, 56);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(168, 40);
			this.groupBox1.TabIndex = 329;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Opciones Valor";
			// 
			// rbDescuento
			// 
			this.rbDescuento.Checked = true;
			this.rbDescuento.Location = new System.Drawing.Point(76, 16);
			this.rbDescuento.Name = "rbDescuento";
			this.rbDescuento.Size = new System.Drawing.Size(84, 16);
			this.rbDescuento.TabIndex = 1;
			this.rbDescuento.TabStop = true;
			this.rbDescuento.Text = "Descuento";
			// 
			// rbFijo
			// 
			this.rbFijo.Location = new System.Drawing.Point(16, 16);
			this.rbFijo.Name = "rbFijo";
			this.rbFijo.Size = new System.Drawing.Size(48, 16);
			this.rbFijo.TabIndex = 0;
			this.rbFijo.Text = "Fijo";
			// 
			// txtValor
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance33;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(680, 104);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(64, 21);
			this.txtValor.TabIndex = 331;
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			this.txtValor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValor_KeyPress);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(624, 106);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(48, 16);
			this.label6.TabIndex = 330;
			this.label6.Text = "VALOR:";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 112);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(142, 14);
			this.label7.TabIndex = 334;
			this.label7.Text = "ARTICULOS DE PROMOCION";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// checkBox1
			// 
			this.checkBox1.Location = new System.Drawing.Point(496, 64);
			this.checkBox1.Name = "checkBox1";
			this.checkBox1.Size = new System.Drawing.Size(128, 16);
			this.checkBox1.TabIndex = 335;
			this.checkBox1.Text = "Todo/Habitaciones";
			this.checkBox1.Click += new System.EventHandler(this.checkBox1_Click);
			// 
			// ugArtPromo
			// 
			this.ugArtPromo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugArtPromo.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugArtPromo.DataSource = this.ultraDataSource4;
			appearance34.BackColor = System.Drawing.Color.White;
			this.ugArtPromo.DisplayLayout.Appearance = appearance34;
			this.ugArtPromo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 60;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 70;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 193;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.VisiblePosition = 3;
			ultraGridColumn26.Width = 300;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 4;
			ultraGridColumn27.Width = 494;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			ultraGridBand5.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ugArtPromo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance35.ForeColor = System.Drawing.Color.Black;
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugArtPromo.DisplayLayout.Override.ActiveRowAppearance = appearance35;
			this.ugArtPromo.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugArtPromo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugArtPromo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance36.BackColor = System.Drawing.Color.Transparent;
			this.ugArtPromo.DisplayLayout.Override.CardAreaAppearance = appearance36;
			this.ugArtPromo.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance37.FontData.BoldAsString = "True";
			appearance37.FontData.Name = "Arial";
			appearance37.FontData.SizeInPoints = 10F;
			appearance37.ForeColor = System.Drawing.Color.White;
			appearance37.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugArtPromo.DisplayLayout.Override.HeaderAppearance = appearance37;
			appearance38.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance38.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArtPromo.DisplayLayout.Override.RowAlternateAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArtPromo.DisplayLayout.Override.RowSelectorAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArtPromo.DisplayLayout.Override.SelectedRowAppearance = appearance40;
			this.ugArtPromo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugArtPromo.Location = new System.Drawing.Point(8, 128);
			this.ugArtPromo.Name = "ugArtPromo";
			this.ugArtPromo.Size = new System.Drawing.Size(832, 96);
			this.ugArtPromo.TabIndex = 336;
			this.ugArtPromo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ugArtPromo.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ugArtPromo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugArtPromo_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// checkBox2
			// 
			this.checkBox2.Checked = true;
			this.checkBox2.CheckState = System.Windows.Forms.CheckState.Checked;
			this.checkBox2.Location = new System.Drawing.Point(496, 80);
			this.checkBox2.Name = "checkBox2";
			this.checkBox2.Size = new System.Drawing.Size(72, 16);
			this.checkBox2.TabIndex = 337;
			this.checkBox2.Text = "Activa";
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Location = new System.Drawing.Point(152, 544);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(136, 32);
			this.button1.TabIndex = 338;
			this.button1.Text = "NUEVO";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// frmPromocionesLote
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(848, 581);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.checkBox2);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNomPromocion);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.ugArtPromo);
			this.Controls.Add(this.checkBox1);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtCant);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.ChkHoras);
			this.Controls.Add(this.ChkDias);
			this.Controls.Add(this.uDiasSemana);
			this.Controls.Add(this.ChkHabitaciones);
			this.Controls.Add(this.uGridHoras);
			this.Controls.Add(this.ugHabitaciones);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.dtFeFin);
			this.Controls.Add(this.dtFeIni);
			this.Controls.Add(this.txtDescripcion);
			this.Name = "frmPromocionesLote";
			this.Text = "Promociones LOTE";
			this.Load += new System.EventHandler(this.frmPromocionesLote_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uDiasSemana)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridHoras)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCant)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugArtPromo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmPromocionesLote_Load(object sender, System.EventArgs e)
		{
			this.dtFeIni.Value = DateTime.Now;
			this.dtFeFin.Value = DateTime.Now;
			string stExecBodega = string.Format("Select idBodega, Bodega, Codigo, Nombre From Bodega Where Activo = 1");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsBodega, stExecBodega);
		}

		private void cdsBodega_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsBodega.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			laBodega = (int)this.cmbBodega.Value;
			string ugSQL = string.Format("Exec Habitacion_Promocion {0},{1}",laBodega.ToString(),idPromo);
			ugHabitaciones.DataSource = Funcion.dvProcedimiento(cdsBodega, ugSQL);
		}

		private void ugHabitaciones_Click(object sender, System.EventArgs e)
		{
			try
			{
				int idHabtcn = (int)this.ugHabitaciones.ActiveRow.Cells["idHabitacion"].Value;
				numHabtc = idHabtcn;
				string hugSQL = string.Format("Exec DiasHabitaciones_Promociones {0},{1}",idHabtcn,idPromo);
				uDiasSemana.DataSource = Funcion.dvProcedimiento(cdsBodega, hugSQL);
				activo = true;
				this.ChkDias.Checked = false;
				activo = false;
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void uDiasSemana_Click(object sender, System.EventArgs e)
		{
			try
			{
				//int idHabtcn = (int)this.uDiasSemana.ActiveRow.Cells["idHabitacion"].Value;
				int idDiaHabitcn = (int)this.uDiasSemana.ActiveRow.Cells["idDiaPromocion"].Value;
				numDia = idDiaHabitcn;
				string hugSQL = string.Format("Exec HorasHabitaciones_Promociones {0},{1}",idDiaHabitcn,idPromo);
				uGridHoras.DataSource = Funcion.dvProcedimiento(cdsBodega, hugSQL);
				actv = true;
				this.ChkHoras.Checked = false;
				actv = false;
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			/* Graba Promoción */
			DateTime dtFecIni = (DateTime) this.dtFeIni.Value;
			DateTime dtFecFin = (DateTime) this.dtFeFin.Value;
			int opVlrs = 0;
			if(this.rbFijo.Checked==true) opVlrs = 1;
			if(this.rbDescuento.Checked==true) opVlrs = 2;
			decimal vlrOp = 0.00m;
			vlrOp = Convert.ToDecimal(this.txtValor.Value);

			using(SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
						
				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Graba Promocion
					#region Maestro
					int sitodo = 0;
					int siactiva = 0;

					if(this.checkBox1.Checked) sitodo = 1;
					if(this.checkBox2.Checked) siactiva = 1;

					string sSQL = string.Format("Exec GrabaPromoLote {0}, {1}, '{2}', '{3}', '{4}', '{5}', {6},{7},{8},{9},{10},{11}", 
						idPromo, laBodega, this.txtNomPromocion.Text.ToString(), 
						dtFecIni.ToString("yyyyMMdd HH:mm"), dtFecFin.ToString("yyyyMMdd HH:mm"), 
						this.txtDescripcion.Text.ToString(), (int)this.txtCant.Value, opVlrs, vlrOp,idArtPromo,siactiva,sitodo);
					oCmdActualiza.CommandText = sSQL;
					idPromo = (int)oCmdActualiza.ExecuteScalar();
					#endregion Maestro
					#region Promcion habitaciones
					if(!this.checkBox1.Checked)
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugHabitaciones.Rows.All)
						{	
							int idHPRo = (int)dr.Cells["idHabitacionPromocion"].Value;
							int nHabtc = (int)dr.Cells["idHabitacion"].Value;
							bool siact = false;
							if ((bool)dr.Cells["Activa"].Value) siact = true;
							else siact = false;
							string sSQLDetalle = string.Format("Exec GrabaPromoHabitaciones {0}, {1}, {2}, {3}",
								idHPRo, idPromo,	nHabtc, siact);
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
							if((bool)dr.Cells["Activa"].Value)
							{
								#region Promcion dias
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drPD in this.uDiasSemana.Rows.All)
								{	
									int idHaPRo = (int)drPD.Cells["idDiaPromocion"].Value;
									int nDiaSem = (int)drPD.Cells["idDia"].Value;
									bool diact = false;
									if ((bool)drPD.Cells["Activa"].Value) diact = true;
									else diact = false;
									string sSQLDePD = string.Format("Exec GrabaDiaPromoHabitacion {0}, {1}, {2}, {3}, {4}",
										idHaPRo, nHabtc, nDiaSem, diact, idPromo);
									oCmdActualiza.CommandText = sSQLDePD;
									oCmdActualiza.ExecuteNonQuery();
									if ((bool)drPD.Cells["Activa"].Value)
									{
										#region Promcion horas
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drPH in this.uGridHoras.Rows.All)
										{	
											int idHrPRo = (int)drPH.Cells["idHoraPromocion"].Value;
											int nHraPro = (int)drPH.Cells["idHora"].Value;
											bool hiact = false;
											if ((bool)drPH.Cells["Activa"].Value) hiact = true;
											else hiact = false;
											string sSQLDePH = string.Format("Exec GrabaHorasPromoHabitacion {0}, {1}, {2}, {3}, {4}",
												idHrPRo, nDiaSem,	nHraPro, hiact, idPromo);
											oCmdActualiza.CommandText = sSQLDePH;
											oCmdActualiza.ExecuteNonQuery();
										}
										#endregion Promcion horas
									}
								}
								#endregion Promcion dias
							}
						}
						#region articulospromo
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drArpr in this.ugArtPromo.Rows.All)
						{
							//PromocionArticulos
							string sSQLap = string.Format("Exec PromocionArticulos {0}, {1}",
								idPromo,
								(int)drArpr.Cells["idArticulo"].Value);
							oCmdActualiza.CommandText = sSQLap;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion articulospromo
					}
					else
					{
						/* Promociones Totales */
						bool siact = false;
						#region Promciondias
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDs in this.uDiasSemana.Rows.All)
						{
							string sSQLDePD = string.Format("Exec GrabaDiaPromoHabitacion {0}, {1}, {2}, {3}, {4}",
								0, 0, (int)drDs.Cells["idDia"].Value, (int)drDs.Cells["Activa"].Value, idPromo);
							oCmdActualiza.CommandText = sSQLDePD;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Promciondias							
						#region articulospromo
						if (this.ugArtPromo.Rows.Count == 0)
						{
							MessageBox.Show("Ingrese los productos para crear PROMOCION", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drArpr in this.ugArtPromo.Rows.All)
						{
							//PromocionArticulos
							string sSQLap = string.Format("Exec PromocionArticulos {0},{1},{2}",
								idPromo,
								(int)drArpr.Cells["idArticulo"].Value,
								(int)drArpr.Cells["idPromoArt"].Value);
							oCmdActualiza.CommandText = sSQLap;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion articulospromo
					}
					#endregion Promcion habitaciones	
					#endregion Graba Promocion

					oTransaction.Commit();

				}
				catch (Exception ex)
				{
					try
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
					MessageBox.Show ("Promoción GRABADA Correctamente", "Aviso", 
						MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				}
			}
		}

		private void txtNomPromocion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaPromocionesLote miBusqdLote = new BuscaPromocionesLote(this.txtNomPromocion.Text.ToString()))
				{
					if (DialogResult.OK == miBusqdLote.ShowDialog())
					{
						this.txtNomPromocion.Text = miBusqdLote.NomProcn;
						idPromo = miBusqdLote.idDPro;
						BuskPromocion(idPromo);
					}
				}
			}

			//if (e.KeyCode == Keys.Enter) BuskPromocion(idPromo);
		}

		public void BuskPromocion(int nPromocn)
		{
			/* Busqueda de promociones */
			string sSQL = string.Format("Exec ConsultaPromocionesLote {0}", nPromocn);
			SqlDataReader drPromocionLt = Funcion.rEscalarSQL(cdsBodega, sSQL, true);
			drPromocionLt.Read();
			if (drPromocionLt.HasRows)
			{
				this.dtFeIni.Value = drPromocionLt.GetDateTime(2);
				this.dtFeFin.Value = drPromocionLt.GetDateTime(3);
				this.txtCant.Value = drPromocionLt.GetInt32(4);
				this.txtDescripcion.Text = drPromocionLt.GetString(5).ToString();
				int tpVlr = drPromocionLt.GetInt32(6);
				if (tpVlr == 1) this.rbFijo.Checked = true;
				else this.rbDescuento.Checked = true;
				this.cmbBodega.Value = drPromocionLt.GetInt32(8);
				this.txtValor.Value = drPromocionLt.GetDecimal(7);
				int tpProm = drPromocionLt.GetInt32(9);
				if (tpProm == 1) this.checkBox1.Checked = true;
			}
			/*Busqueda articulos promociones*/
			string QryDetFc = string.Format("Exec CargaArtPromocion {0}", nPromocn);
			ugArtPromo.DataSource = Funcion.dvProcedimiento(cdsBodega, QryDetFc);

			/*Busqueda de dias promociones*/
			string QryDP = string.Format("Exec CargaDiasPromocion {0}", nPromocn);
			uDiasSemana.DataSource = Funcion.dvProcedimiento(cdsBodega, QryDP);
		}

		private void txtNomPromocion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbBodega.Focus();
		}

		private void cmbBodega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFeIni.Focus();
		}

		private void dtFeFin_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCant.Focus();
		}

		private void txtCant_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		  if (e.KeyChar == 13) this.btnGuardar.Focus();
		}

		private void ChkHabitaciones_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.ChkHabitaciones.Checked)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugHabitaciones.Rows.All)
				{
					dr.Cells["Activa"].Value = true;
				}
			}
			else
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugHabitaciones.Rows.All)
				{
					dr.Cells["Activa"].Value = false;
				}
			}
		}

		private void ChkDias_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.ChkDias.Checked)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uDiasSemana.Rows.All)
				{
					dr.Cells["Activa"].Value = true;
				}
			}
			else
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uDiasSemana.Rows.All)
				{
					dr.Cells["Activa"].Value = false;
				}
			}
		}

		private void ChkHoras_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.ChkHoras.Checked)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridHoras.Rows.All)
				{
					dr.Cells["Activa"].Value = true;
				}
			}
			else
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridHoras.Rows.All)
				{
					dr.Cells["Activa"].Value = false;
				}
			}
		}

		private void txtValor_Enter(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
		  this.txtValor.SelectAll();
		}

		private void txtValor_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCant.Focus();
		}

		private void txtCant_Click(object sender, System.EventArgs e)
		{
		  this.txtCant.SelectAll();
		}

		private void txtCant_Enter(object sender, System.EventArgs e)
		{
			this.txtCant.SelectAll();
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			int nBodega = 0;
			DateTime FechaPromo = DateTime.Now;

			if (this.cmbBodega.ActiveRow != null) nBodega = (int)this.cmbBodega.Value;
			if (e.KeyCode == Keys.F3)
			{
				if (this.ugArtPromo.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO")
				{
					//if (!FuncionesInventario.vSerialesIngresadasGridAbreInventario(this.ultraGrid1, this.grdSerial, bNuevo)) return;
					//FuncionesInventario.AbreBusquedaArticulos(this.ugArtPromo, nBodega, 1, 1, Convert.ToDateTime(this.dtFeIni.Value));
					using (BuscaArticulos Busqueda = new BuscaArticulos(this.ugArtPromo.ActiveRow.Cells["Codigo"].Value.ToString(), nBodega, 1, 1, Convert.ToDateTime(this.dtFeIni.Value), false))
					{
						if (DialogResult.OK == Busqueda.ShowDialog())
						{									
							this.ugArtPromo.ActiveRow.Cells["Codigo"].Value = Busqueda.sCodigo;
							this.ugArtPromo.ActiveRow.Cells["Articulo"].Value = Busqueda.sArticulo;
							this.ugArtPromo.ActiveRow.Cells["idArticulo"].Value = Busqueda.idArticulo;
							this.ugArtPromo.ActiveRow.Cells["idPromoArt"].Value = 0;
							this.ugArtPromo.ActiveRow.Cells["idPromocion"].Value = 0;
						}
					}	
				}
			}
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ugArtPromo);
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Codigo
//			if (e.Cell.Column.ToString() == "Codigo")
//			{
//				if (!bModificaColor)
//					FuncionesInventario.FacturacionCargaArticulos(sender, e, this.ultraGrid1, (int)this.cmbFormaPago.Value, idTipoPrecio, (int)this.cmbBodega.Value, 
//						(int)this.txtidBanco.Value, (int)this.txtidTarjeta.Value, (int)this.txtidPlan.Value, (int)this.txtidPlazo.Value, Convert.ToDateTime(this.dtFecha.Value), 1, 
//						true, bPromociones, bPromociones, bPromociones, this.NumCuotas, (int)this.txtIdEF.Value, (int)this.txtNumIdCliente.Value, iNumCompras, 0, cdsSeteoF);
//				
//				if (this.ultraGrid1.Rows.Count > 0) if ((int)this.cmbBodega.Value == 69) this.txtClienteNombre.Enabled = false;
//
//				bActualiza = true;
//				this.Total();	
//			}		
			#endregion Codigo		
		}

		private void checkBox1_Click(object sender, System.EventArgs e)
		{
			/*Visuliza dias de semana*/
			try
			{
				if(this.checkBox1.Checked)
				{
					string hugSQL = string.Format("Exec PromocionDias 1");
					uDiasSemana.DataSource = Funcion.dvProcedimiento(cdsBodega, hugSQL);
				}
				else 
				{
					string hugSQL = string.Format("Exec PromocionDias 0");
					uDiasSemana.DataSource = Funcion.dvProcedimiento(cdsBodega, hugSQL);
				}
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void ugArtPromo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			/*Inicializar datos del GRID*/
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			/*Encerar variables*/
			this.cmbBodega.Value = 0;
			this.txtNomPromocion.Text = "";
			this.txtDescripcion.Text = "";
			this.checkBox1.Checked = false;
			this.checkBox2.Checked = true;
			this.txtValor.Value = 0;
			this.txtCant.Value = 0;
			FuncionesProcedimientos.EstadoGrids(false, this.ugArtPromo);
			FuncionesProcedimientos.EstadoGrids(false, this.uDiasSemana);
			FuncionesProcedimientos.EstadoGrids(false, this.ugHabitaciones);			
		}

		private void ugHabitaciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
		
	}
}
