using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_ResumenDeComisionesSemanal.
	/// </summary>
	public class frmCom_ResumenDeComisionesSemanal : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridVendedores;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optConsulta;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocales;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optRegiones;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCargos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_ResumenDeComisionesSemanal()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_ResumenDeComisionesSemanal));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_AsignacionDeVendedoresSemanal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaSemanal");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaContado");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaCredito");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaGEx");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaContado");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaCredito");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaCarpa");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaGEx");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaTotal");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeDiario");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeSemanal");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeContado");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCredito");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeGEx");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionContado");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionCredito");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionCarpa");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionGEx");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estrella1", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estrella2", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estrella3", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estrella4", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estrella5", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estrella6", 5);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_AsignacionDeVendedoresSemanal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaSemanal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaContado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaContado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeDiario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeSemanal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeContado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionContado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contador");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			this.btnVer = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridVendedores = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.optConsulta = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnExcel = new System.Windows.Forms.Button();
			this.optRegiones = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridCargos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.SuspendLayout();
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(368, 7);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 733;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(192, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 17);
			this.label4.TabIndex = 730;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 10);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(35, 17);
			this.label5.TabIndex = 729;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 728;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 727;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtDesde_BeforeDropDown);
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// uGridVendedores
			// 
			this.uGridVendedores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVendedores.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVendedores.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridVendedores.DisplayLayout.Appearance = appearance3;
			ultraGridBand1.AddButtonCaption = "Serie";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 190;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 200;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 220;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Meta Semanal";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 95;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Meta Contado";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 95;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Meta Credito";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 95;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance7;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Meta GEx";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 95;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance8;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Venta Contado";
			ultraGridColumn10.Header.VisiblePosition = 14;
			ultraGridColumn10.Width = 95;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance9;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Venta Credito";
			ultraGridColumn11.Header.VisiblePosition = 15;
			ultraGridColumn11.Width = 95;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance10;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Venta Carpa";
			ultraGridColumn12.Header.VisiblePosition = 16;
			ultraGridColumn12.Width = 95;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance11;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Venta GEx";
			ultraGridColumn13.Header.VisiblePosition = 17;
			ultraGridColumn13.Width = 95;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance12;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Venta Total";
			ultraGridColumn14.Header.VisiblePosition = 18;
			ultraGridColumn14.Width = 95;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance13;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "% Diario";
			ultraGridColumn15.Header.VisiblePosition = 9;
			ultraGridColumn15.Width = 95;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance14;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "% Semanal";
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Width = 95;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance15;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "% Contado";
			ultraGridColumn17.Header.VisiblePosition = 11;
			ultraGridColumn17.Width = 95;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance16;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "% Credito";
			ultraGridColumn18.Header.VisiblePosition = 12;
			ultraGridColumn18.Width = 95;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance17;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "% GEx";
			ultraGridColumn19.Header.VisiblePosition = 13;
			ultraGridColumn19.Width = 95;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance18;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "Com Contado";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 95;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance19;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Com Credito";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 95;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance20;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "Com Carpa";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 95;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance21;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "Com GEx";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 95;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance22;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "Total Comisión";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 100;
			ultraGridColumn25.Header.VisiblePosition = 25;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn26.Header.Caption = "";
			ultraGridColumn26.Header.VisiblePosition = 24;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn26.Width = 15;
			ultraGridColumn27.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn27.Header.Caption = "";
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn27.Width = 15;
			ultraGridColumn28.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn28.Header.Caption = "";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn28.Width = 15;
			ultraGridColumn29.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn29.Header.Caption = "";
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn29.Width = 15;
			ultraGridColumn30.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn30.Header.Caption = "";
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn30.Width = 15;
			ultraGridColumn31.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn31.Header.Caption = "";
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn31.Width = 15;
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
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridVendedores.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVendedores.DisplayLayout.Override.ActiveRowAppearance = appearance23;
			this.uGridVendedores.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVendedores.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVendedores.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.uGridVendedores.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance25.FontData.BoldAsString = "True";
			appearance25.FontData.Name = "Arial";
			appearance25.FontData.SizeInPoints = 8.5F;
			appearance25.ForeColor = System.Drawing.Color.White;
			appearance25.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVendedores.DisplayLayout.Override.HeaderAppearance = appearance25;
			this.uGridVendedores.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVendedores.DisplayLayout.Override.RowAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVendedores.DisplayLayout.Override.RowSelectorAppearance = appearance27;
			this.uGridVendedores.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVendedores.Location = new System.Drawing.Point(256, 56);
			this.uGridVendedores.Name = "uGridVendedores";
			this.uGridVendedores.Size = new System.Drawing.Size(1100, 504);
			this.uGridVendedores.TabIndex = 736;
			this.uGridVendedores.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridVendedores_InitializeRow);
			this.uGridVendedores.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVendedores_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
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
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn18.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn19.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn20.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn21.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn22.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn23.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn25.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn26.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn27.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn26,
																																 ultraDataColumn27});
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
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(952, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 737;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1364, 8);
			this.groupBox1.TabIndex = 738;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// optConsulta
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.optConsulta.Appearance = appearance28;
			this.optConsulta.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optConsulta.CheckedIndex = 0;
			this.optConsulta.Enabled = false;
			this.optConsulta.ItemAppearance = appearance29;
			this.optConsulta.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = false;
			valueListItem1.DisplayText = "Resumen Por Vendedor";
			valueListItem2.DataValue = true;
			valueListItem2.DisplayText = "Resumen Por Local";
			this.optConsulta.Items.Add(valueListItem1);
			this.optConsulta.Items.Add(valueListItem2);
			this.optConsulta.ItemSpacingVertical = 10;
			this.optConsulta.Location = new System.Drawing.Point(456, 6);
			this.optConsulta.Name = "optConsulta";
			this.optConsulta.Size = new System.Drawing.Size(288, 24);
			this.optConsulta.TabIndex = 739;
			this.optConsulta.Text = "Resumen Por Vendedor";
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.CausesValidation = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1244, 7);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(112, 23);
			this.btnExcel.TabIndex = 740;
			this.btnExcel.Text = "&Exportar a Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// optRegiones
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.optRegiones.Appearance = appearance30;
			this.optRegiones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optRegiones.CheckedIndex = 0;
			this.optRegiones.Enabled = false;
			this.optRegiones.ItemAppearance = appearance31;
			this.optRegiones.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "Todos";
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "Sierra";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "Costa";
			this.optRegiones.Items.Add(valueListItem3);
			this.optRegiones.Items.Add(valueListItem4);
			this.optRegiones.Items.Add(valueListItem5);
			this.optRegiones.ItemSpacingVertical = 10;
			this.optRegiones.Location = new System.Drawing.Point(752, 6);
			this.optRegiones.Name = "optRegiones";
			this.optRegiones.Size = new System.Drawing.Size(184, 24);
			this.optRegiones.TabIndex = 741;
			this.optRegiones.Text = "Todos";
			this.optRegiones.ValueChanged += new System.EventHandler(this.optRegiones_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(bool);
			ultraDataColumn30.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30});
			// 
			// uGridLocales
			// 
			this.uGridLocales.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocales.DataSource = this.ultraDataSource2;
			appearance32.BackColor = System.Drawing.Color.White;
			this.uGridLocales.DisplayLayout.Appearance = appearance32;
			this.uGridLocales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 100;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridColumn33.Width = 199;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn34.Header.Caption = "...";
			ultraGridColumn34.Header.VisiblePosition = 2;
			ultraGridColumn34.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn34.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			ultraGridBand2.Header.Caption = "LOCALES";
			ultraGridBand2.HeaderVisible = true;
			this.uGridLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance33.ForeColor = System.Drawing.Color.Black;
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocales.DisplayLayout.Override.ActiveRowAppearance = appearance33;
			this.uGridLocales.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocales.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLocales.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance34.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocales.DisplayLayout.Override.CardAreaAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance35.FontData.BoldAsString = "True";
			appearance35.FontData.Name = "Arial";
			appearance35.FontData.SizeInPoints = 8.5F;
			appearance35.ForeColor = System.Drawing.Color.White;
			appearance35.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocales.DisplayLayout.Override.HeaderAppearance = appearance35;
			appearance36.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance36.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowAlternateAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowSelectorAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.SelectedRowAppearance = appearance38;
			this.uGridLocales.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales.Location = new System.Drawing.Point(8, 224);
			this.uGridLocales.Name = "uGridLocales";
			this.uGridLocales.Size = new System.Drawing.Size(240, 336);
			this.uGridLocales.TabIndex = 742;
			this.uGridLocales.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridLocales_InitializeLayout);
			// 
			// uGridCargos
			// 
			this.uGridCargos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCargos.DataSource = this.ultraDataSource4;
			appearance39.BackColor = System.Drawing.Color.White;
			this.uGridCargos.DisplayLayout.Appearance = appearance39;
			this.uGridCargos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 73;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Width = 199;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn37.Header.Caption = "...";
			ultraGridColumn37.Header.VisiblePosition = 2;
			ultraGridColumn37.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn37.Width = 20;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			ultraGridBand3.Header.Caption = "CARGOS";
			ultraGridBand3.HeaderVisible = true;
			this.uGridCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance40.ForeColor = System.Drawing.Color.Black;
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCargos.DisplayLayout.Override.ActiveRowAppearance = appearance40;
			this.uGridCargos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCargos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCargos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance41.BackColor = System.Drawing.Color.Transparent;
			this.uGridCargos.DisplayLayout.Override.CardAreaAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance42.FontData.BoldAsString = "True";
			appearance42.FontData.Name = "Arial";
			appearance42.FontData.SizeInPoints = 8.5F;
			appearance42.ForeColor = System.Drawing.Color.White;
			appearance42.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCargos.DisplayLayout.Override.HeaderAppearance = appearance42;
			appearance43.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance43.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargos.DisplayLayout.Override.RowAlternateAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargos.DisplayLayout.Override.RowSelectorAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargos.DisplayLayout.Override.SelectedRowAppearance = appearance45;
			this.uGridCargos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCargos.Location = new System.Drawing.Point(8, 56);
			this.uGridCargos.Name = "uGridCargos";
			this.uGridCargos.Size = new System.Drawing.Size(240, 160);
			this.uGridCargos.TabIndex = 743;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn31.DefaultValue = 0;
			ultraDataColumn33.DataType = typeof(bool);
			ultraDataColumn33.DefaultValue = false;
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33});
			// 
			// frmCom_ResumenDeComisionesSemanal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1364, 566);
			this.Controls.Add(this.uGridCargos);
			this.Controls.Add(this.optRegiones);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.optConsulta);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.uGridVendedores);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridLocales);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmCom_ResumenDeComisionesSemanal";
			this.Text = "Resumen De Comisiones Semanal";
			this.Load += new System.EventHandler(this.frmCom_ResumenDeComisionesSemanal_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		int iSemana = 0;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCom_ResumenDeComisionesSemanal_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0720");
			
			if (!Funcion.Permiso("658", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Resumen de Comisiones", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				

				return;
			}
			
			FuncionesProcedimientos.EstadoGrids(false, this.uGridCargos);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridLocales);

			#region Accesos
			if (miAcceso.BExportar) this.btnExcel.Visible = true;
			if (miAcceso.BCom_FiltrarPorCargo) FuncionesProcedimientos.EstadoGrids(true, this.uGridCargos);

//			if (miAcceso.BCom_CerrarPeriodo)
//			{
//				this.btnBloquear.Visible = true;
//				this.btnCerrar.Visible = true;
//				this.btnImprimir.Visible = true;
//			}

//			if (miAcceso.BCom_Observaciones)
//			{
//				this.uGridVendedores.DisplayLayout.Bands[0].Columns["btnObservaciones"].Hidden = false;
//				this.uGridVendedores.DisplayLayout.Bands[0].Columns["btnResumen"].Hidden = false;
//				this.uGridVendedores.DisplayLayout.Bands[0].Columns["btnImprimir"].Hidden = false;
//				//this.uGridVendedores.DisplayLayout.Bands[0].Columns["R"].Hidden = false;
//			}

//			if (miAcceso.BVerComisiones)
//			{
//				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Comision"].Hidden = false;
//				//					this.uGridVendedores.DisplayLayout.Bands[0].Columns["Com_Carpa"].Hidden = false;
//				//					this.uGridVendedores.DisplayLayout.Bands[0].Columns["BonoCarpa"].Hidden = false;
//				this.uGridVendedores.DisplayLayout.Bands[0].Columns["BonoCredito"].Hidden = false;
//				this.uGridVendedores.DisplayLayout.Bands[0].Columns["ComisionGE"].Hidden = false;
//				this.uGridVendedores.DisplayLayout.Bands[0].Columns["ComisionTotal"].Hidden = false;
//			}
			#endregion Accesos
						
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("02/07/2018");

			DateTime dtDesde = DateTime.Today;

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Exec Com_RetornaPeriodosSemanales", true);
			dr.Read();
			if (dr.HasRows)
			{		
				dtDesde = dr.GetDateTime(0);
				this.dtHasta.Value = dr.GetDateTime(1);
			}
			dr.Close();

			this.dtDesde.Value = dtDesde;

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);		

			if (iBodega == 0) FuncionesProcedimientos.EstadoGrids(true, this.uGridLocales);

			this.optRegiones.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Com_ResumenSemanalValidaRegion {0}, {1}", MenuLatinium.IdUsuario, MenuLatinium.iNivel));
				
			if (MenuLatinium.iNivel != 23 && MenuLatinium.iNivel != 25) this.optConsulta.Enabled = true;

			Cargos();

			if ((int)this.optRegiones.Value == 0) this.optRegiones.Enabled = true;
			//this.dtHasta.CalendarInfo.MinDate = Convert.ToDateTime(this.dtDesde.Value);
			this.dtHasta.Value = DateTime.Today;			
			//		this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;
			//dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sLocales = "";
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString();
					else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			if (sLocales.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			string sCargos = "";

			if (!(bool)this.optConsulta.Value)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCargos.Rows.All)
				{
					if ((bool) dr.Cells["Sel"].Value)
					{
						if (sCargos.Length == 0) sCargos = dr.Cells["idCargo"].Value.ToString();
						else if (sCargos.Length > 0) sCargos = sCargos + ", " + dr.Cells["idCargo"].Value.ToString();
					}
				}

				if (sCargos.Length == 0)
				{
					MessageBox.Show("Seleccione al menos un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}

			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			#endregion Valida Fechas

			string sSQL = "";

			if (!(bool)this.optConsulta.Value)
			{
				sSQL = string.Format("Exec Com_ResumenDeComisionesSemanal '{0}', '{1}', '{2}', '{3}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, sCargos);
			}
			else
			{
				sSQL = string.Format("Exec Com_ResumenSemanalPorLocales '{0}', '{1}', {2}, '{3}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), iSemana, sLocales);
			}

			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridVendedores);

			this.lblContador.Text = this.uGridVendedores.Rows.Count + " Registros Encontrados";
		}

		private void Cargos()
		{
			this.uGridCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Com_ResumenComisionesListaCargos {0}, {1}, {2}", 
				MenuLatinium.IdUsuario, MenuLatinium.iNivel, (int)this.optRegiones.Value));
		}

		private void Locales()
		{
			string sSQL = string.Format("Exec Com_ResumenComisionesListaLocales {0}, {1}, {2}, '{3}'", 
				MenuLatinium.IdUsuario, iSemana, (int)this.optRegiones.Value, Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridLocales);
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value != System.DBNull.Value)
			{
				this.dtHasta.Value = Convert.ToDateTime(this.dtDesde.Value).AddDays(6);

				iSemana = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Semana From Com_PresupuestoSemanal Where CONVERT(Date, Desde) = '{0}' And CONVERT(Date, Hasta) = '{1}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));

				Locales();
			}
		}

		private void uGridVendedores_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (e.Row.Cells["idCargo"].Value != System.DBNull.Value)
			{
				if ((int)e.Row.Cells["idCargo"].Value != 4 && (int)e.Row.Cells["idCargo"].Value != 5 && (int)e.Row.Cells["idCargo"].Value != 7)
				{
					if ((int)e.Row.Cells["Contador"].Value == 1) e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png");
					if ((int)e.Row.Cells["Contador"].Value == 2)
					{
						e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella2"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
					}
					if ((int)e.Row.Cells["Contador"].Value == 3) 
					{
						e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella2"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella3"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
					}
					if ((int)e.Row.Cells["Contador"].Value == 4) 
					{
						e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella2"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella3"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella4"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
					}
					if ((int)e.Row.Cells["Contador"].Value == 5) 
					{
						e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella2"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella3"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella4"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella5"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
					}
					if ((int)e.Row.Cells["Contador"].Value == 6) 
					{
						e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella2"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella3"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella4"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella5"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
						e.Row.Cells["Estrella6"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
					}
			
					if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 100.00m) e.Row.Cells["PorcentajeDiario"].Appearance.BackColor = Color.LimeGreen;
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 100.00m) e.Row.Cells["PorcentajeDiario"].Appearance.BackColor = Color.Yellow;
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 85.00m) e.Row.Cells["PorcentajeDiario"].Appearance.BackColor = Color.Orange;			
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 70.00m) e.Row.Cells["PorcentajeDiario"].Appearance.BackColor = Color.Red;

					if (Convert.ToDecimal(e.Row.Cells["PorcentajeSemanal"].Value) >= 100.00m) e.Row.Cells["PorcentajeSemanal"].Appearance.BackColor = Color.LimeGreen;
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeSemanal"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeSemanal"].Value) < 100.00m) e.Row.Cells["PorcentajeSemanal"].Appearance.BackColor = Color.Yellow;
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeSemanal"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeSemanal"].Value) < 85.00m) e.Row.Cells["PorcentajeSemanal"].Appearance.BackColor = Color.Orange;			
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeSemanal"].Value) < 70.00m) e.Row.Cells["PorcentajeSemanal"].Appearance.BackColor = Color.Red;

					if (Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) >= 100.00m) e.Row.Cells["PorcentajeContado"].Appearance.BackColor = Color.LimeGreen;
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) < 100.00m) e.Row.Cells["PorcentajeContado"].Appearance.BackColor = Color.Yellow;
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) < 85.00m) e.Row.Cells["PorcentajeContado"].Appearance.BackColor = Color.Orange;			
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) < 70.00m) e.Row.Cells["PorcentajeContado"].Appearance.BackColor = Color.Red;

					if (Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) >= 100.00m) e.Row.Cells["PorcentajeCredito"].Appearance.BackColor = Color.LimeGreen;
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) < 100.00m) e.Row.Cells["PorcentajeCredito"].Appearance.BackColor = Color.Yellow;
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) < 85.00m) e.Row.Cells["PorcentajeCredito"].Appearance.BackColor = Color.Orange;			
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) < 70.00m) e.Row.Cells["PorcentajeCredito"].Appearance.BackColor = Color.Red;

					if (Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) >= 100.00m) e.Row.Cells["PorcentajeGEx"].Appearance.BackColor = Color.LimeGreen;
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) < 100.00m) e.Row.Cells["PorcentajeGEx"].Appearance.BackColor = Color.Yellow;
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) < 85.00m) e.Row.Cells["PorcentajeGEx"].Appearance.BackColor = Color.Orange;			
					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) < 70.00m) e.Row.Cells["PorcentajeGEx"].Appearance.BackColor = Color.Red;
				}
			}
		}

		private void uGridVendedores_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridVendedores);
		}

		private void dtDesde_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void optRegiones_ValueChanged(object sender, System.EventArgs e)
		{
			Cargos();

			Locales();
		}

		private void uGridLocales_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}
	}
}
