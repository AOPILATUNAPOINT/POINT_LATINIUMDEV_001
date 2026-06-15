using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPlanificacionActividadesPersonal.
	/// </summary>
	public class frmPlanificacionActividadesPersonal : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnSalir;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNomina;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPlanificacionActividadesPersonal()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActividadesPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaResponsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAsignacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Inicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fin");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimaRevision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesUltRev");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Finaliza");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesFin");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregable");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEditar", 0);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPlanificacionActividadesPersonal));
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeguimiento", 1);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnFinalizar", 2);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDescripcion", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnObservacionesUltRev", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnObservacionesFin", 5);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActividadesPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaResponsable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAsignacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Inicio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fin");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaUltimaRevision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesUltRev");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Finaliza");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesFin");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNomina = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNomina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 95;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 98;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 105;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 200;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Fec Asignacion";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 75;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 75;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance2;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 35;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.CellMultiLine = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 400;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Fecha Seguim";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn11.CellMultiLine = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn11.Header.Caption = "Observaciones Seguimiento";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 180;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 13;
			ultraGridColumn12.Width = 75;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn13.CellMultiLine = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn13.Header.Caption = "Observaciones Finalización";
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 180;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance3;
			ultraGridColumn14.Header.VisiblePosition = 16;
			ultraGridColumn14.Width = 70;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 17;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 55;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance4;
			ultraGridColumn16.Header.Caption = "Estado";
			ultraGridColumn16.Header.VisiblePosition = 18;
			ultraGridColumn16.Width = 90;
			ultraGridColumn17.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance5.Image = ((object)(resources.GetObject("appearance5.Image")));
			appearance5.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance5.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn17.CellButtonAppearance = appearance5;
			ultraGridColumn17.Header.Caption = "E";
			ultraGridColumn17.Header.VisiblePosition = 19;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn17.Width = 30;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			appearance6.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance6.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn18.CellButtonAppearance = appearance6;
			ultraGridColumn18.Header.Caption = "S";
			ultraGridColumn18.Header.VisiblePosition = 20;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 30;
			ultraGridColumn19.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance7.Image = ((object)(resources.GetObject("appearance7.Image")));
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance7.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn19.CellButtonAppearance = appearance7;
			ultraGridColumn19.Header.Caption = "F";
			ultraGridColumn19.Header.VisiblePosition = 21;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 30;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn20.Header.Caption = "...";
			ultraGridColumn20.Header.VisiblePosition = 9;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn20.Width = 20;
			ultraGridColumn21.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn21.Header.Caption = "...";
			ultraGridColumn21.Header.VisiblePosition = 12;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn21.Width = 20;
			ultraGridColumn22.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn22.Header.Caption = "...";
			ultraGridColumn22.Header.VisiblePosition = 15;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn22.Width = 20;
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
																										 ultraGridColumn22});
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridInformacion.DisplayLayout.GroupByBox.Prompt = " ";
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8.5F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance11.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance11.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowAlternateAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.uGridInformacion.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(0, 88);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1144, 240);
			this.uGridInformacion.TabIndex = 4;
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn6.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn8.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn10.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn12.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn15.DataType = typeof(int);
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
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(1046, 7);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(88, 23);
			this.btnNuevo.TabIndex = 18;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1046, 39);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(88, 23);
			this.btnSalir.TabIndex = 22;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 42);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 16);
			this.label1.TabIndex = 641;
			this.label1.Text = "Realiza";
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(416, 38);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 643;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// cmbNomina
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbNomina.Appearance = appearance14;
			this.cmbNomina.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNomina.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNomina.DataSource = this.ultraDataSource2;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn25.Header.VisiblePosition = 2;
			ultraGridColumn25.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbNomina.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNomina.DisplayMember = "Nombre";
			this.cmbNomina.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNomina.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNomina.Location = new System.Drawing.Point(88, 40);
			this.cmbNomina.Name = "cmbNomina";
			this.cmbNomina.Size = new System.Drawing.Size(280, 21);
			this.cmbNomina.TabIndex = 644;
			this.cmbNomina.ValueMember = "idNomina";
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
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1142, 8);
			this.groupBox1.TabIndex = 645;
			this.groupBox1.TabStop = false;
			// 
			// optEstados
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance15;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 1;
			this.optEstados.ItemAppearance = appearance16;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendientes";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Atrasados";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Terminados";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(416, 6);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(328, 24);
			this.optEstados.TabIndex = 646;
			this.optEstados.Text = "Pendientes";
			// 
			// cmbResponsable
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance17;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResponsable.DataSource = this.ultraDataSource2;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Width = 300;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(88, 8);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.Size = new System.Drawing.Size(280, 21);
			this.cmbResponsable.TabIndex = 649;
			this.cmbResponsable.ValueMember = "idNomina";
			this.cmbResponsable.ValueChanged += new System.EventHandler(this.cmbResponsable_ValueChanged);
			this.cmbResponsable.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbResponsable_InitializeLayout);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(70, 16);
			this.label2.TabIndex = 648;
			this.label2.Text = "Responsable";
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(512, 42);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 650;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(856, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(50, 16);
			this.label3.TabIndex = 651;
			this.label3.Text = "E: Editar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(856, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(84, 16);
			this.label4.TabIndex = 652;
			this.label4.Text = "S: Seguimiento";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmPlanificacionActividadesPersonal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1142, 326);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.cmbResponsable);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbNomina);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.uGridInformacion);
			this.KeyPreview = true;
			this.Name = "frmPlanificacionActividadesPersonal";
			this.Text = "Planificación Actividades Personal";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPlanificacionActividadesPersonal_KeyDown);
			this.Load += new System.EventHandler(this.frmPlanificacionActividadesPersonal_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNomina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmPlanificacionActividadesPersonal_Load(object sender, System.EventArgs e)
		{
			this.cmbResponsable.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec APListaResponsables {0}", MenuLatinium.IdUsuario));
			this.cmbResponsable.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario Where idUsuario = {0})) Select idNomina From Nomina Where Codigo In (Select Nombre From Usuario Where idUsuario = {0}) Else Select 0", MenuLatinium.IdUsuario));
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void Registro(int idActividadesPersonal, DateTime dtFecha)
		{
			if (!Validacion.vbComboVacio(this.cmbNomina, "Seleccione el nombre de la persona a Asignar el Trabajo")) return;

			using (frmGuardaActividadesPersonal GAP = new frmGuardaActividadesPersonal(idActividadesPersonal, dtFecha,true))
			{
				
				if (DialogResult.OK == GAP.ShowDialog())
				{
					DateTime dtFecAsigna = DateTime.Today;

					string sSQL = string.Format("Exec GuardaActividadesPersonal {0}, {1}, {2}, '{3}', '{4}', {5}, '{6}', '{7}', {8}", 
						(int)GAP.txtIdActividadesPersonal.Value, (int)this.cmbResponsable.Value, (int)this.cmbNomina.Value,
						dtFecAsigna.ToString("yyyyMMdd"), 
						Convert.ToDateTime(GAP.dtpInicio.Value).ToString("yyyyMMdd"), (int)GAP.txtDias.Value, 
						Convert.ToDateTime(GAP.dtpFinaliza.Value).ToString("yyyyMMdd"),
						GAP.txtDescripcion.Text.ToString(), GAP.chkEntregable.Checked);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					Consulta();	
				}
			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec APValidaResponsableActividad {0}, {1}", 
				MenuLatinium.IdUsuario, (int)this.cmbResponsable.Value), true))
			{
				MessageBox.Show("No puede Asignar una tarea a nombre de otro usuario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			Registro(0, DateTime.Today);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{	
			if (e.Cell.Column.ToString() == "btnDescripcion")			
				FuncionesProcedimientos.AbreVistaMensajes(e.Cell.Row.Cells["Descripcion"].Value.ToString());

			if (e.Cell.Column.ToString() == "btnObservacionesUltRev")	
				FuncionesProcedimientos.AbreVistaMensajes(e.Cell.Row.Cells["ObservacionesUltRev"].Value.ToString());

			if (e.Cell.Column.ToString() == "btnObservacionesFin")	
				FuncionesProcedimientos.AbreVistaMensajes(e.Cell.Row.Cells["ObservacionesFin"].Value.ToString());
		
			if (e.Cell.Column.ToString() == "btnEditar")
			{
				if (e.Cell.Row.Cells["idActividadesPersonal"].Value == System.DBNull.Value) return;

				if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec APValidaResponsableActividad {0}, {1}", 
					MenuLatinium.IdUsuario, (int)this.cmbResponsable.Value), true))
				{
					MessageBox.Show("No puede editar una tarea asignada por otro usuario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				Registro((int)e.Cell.Row.Cells["idActividadesPersonal"].Value, Convert.ToDateTime(e.Cell.Row.Cells["FechaAsignacion"].Value));
			}
      bool bSeguimiento = true;
			if (e.Cell.Column.ToString() == "btnSeguimiento")
			{
				if (e.Cell.Row.Cells["idActividadesPersonal"].Value == System.DBNull.Value) return;

				DateTime dtFecha = Convert.ToDateTime(e.Cell.Row.Cells["Inicio"].Value);

				if (e.Cell.Row.Cells["FechaUltimaRevision"].Value != System.DBNull.Value) 
					dtFecha = Convert.ToDateTime(e.Cell.Row.Cells["FechaUltimaRevision"].Value);

				if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec APValidaResponsableActividad {0}, {1}", 
					MenuLatinium.IdUsuario, (int)this.cmbResponsable.Value), true))
				{
          bSeguimiento = false;
				}
				using (frmAPHistorial APH = new frmAPHistorial((int)e.Cell.Row.Cells["idNomina"].Value, 
								 (int)e.Cell.Row.Cells["idActividadesPersonal"].Value, 
								 e.Cell.Row.Cells["Descripcion"].Value.ToString(), true, bSeguimiento))
				{
					if (DialogResult.OK == APH.ShowDialog())
					{
						Consulta();
					}
				}
				
//				using (frmSeguimientoActividadesPersonal SAP = new frmSeguimientoActividadesPersonal(false, dtFecha))
//				{
//					if (DialogResult.OK == SAP.ShowDialog())
//					{
//						string sSQL = string.Format("Exec GuardaHistorialRevisionActividadesPersonal {0}, '{1}', '{2}'", 
//							(int)e.Cell.Row.Cells["idActividadesPersonal"].Value, Convert.ToDateTime(SAP.dtpFecha.Value).ToString("yyyyMMdd"), SAP.txtDescripcion.Text.ToString());
//						Funcion.EjecutaSQL(cdsSeteoF, sSQL);
//
//						Consulta();
//					}
//				}
			}

//			if (e.Cell.Column.ToString() == "btnFinalizar")
//			{
//				if (e.Cell.Row.Cells["idActividadesPersonal"].Value == System.DBNull.Value) return;
//
//				DateTime dtFecha = Convert.ToDateTime(e.Cell.Row.Cells["Inicio"].Value);
//
//				if (e.Cell.Row.Cells["FechaUltimaRevision"].Value != System.DBNull.Value) 
//					dtFecha = Convert.ToDateTime(e.Cell.Row.Cells["FechaUltimaRevision"].Value);
//
//				using (frmSeguimientoActividadesPersonal SAP = new frmSeguimientoActividadesPersonal(true, dtFecha))
//				{
//					if (DialogResult.OK == SAP.ShowDialog())
//					{
//						string sSQL = string.Format("Exec FinalizaActividadesPersonal '{0}', '{1}', {2}", 
//							Convert.ToDateTime(SAP.dtpFecha.Value).ToString("yyyyMMdd"), SAP.txtDescripcion.Text.ToString(), (int)e.Cell.Row.Cells["idActividadesPersonal"].Value);
//						Funcion.EjecutaSQL(cdsSeteoF, sSQL);
//
//						Consulta();
//					}
//				}
//			}
		}

		private void Consulta ()
		{
			int idNominaResponsable = 0;
			int idNomina = 0;
			
			if (this.cmbResponsable.ActiveRow != null) idNominaResponsable = (int)this.cmbResponsable.Value;
			if (this.cmbNomina.ActiveRow != null) idNomina = (int)this.cmbNomina.Value;

			string sSQL = string.Format("Exec ConsultaActividadesPersonal {0}, {1}, {2}", idNominaResponsable, idNomina, (int)this.optEstados.Value);
			this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			this.uGridInformacion.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);			
			
			if (this.cmbNomina.ActiveRow != null) this.uGridInformacion.Rows.ExpandAll(true); else this.uGridInformacion.Rows.CollapseAll(true);

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Consulta();	
		}

		private void frmPlanificacionActividadesPersonal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cmbResponsable_ValueChanged(object sender, System.EventArgs e)
		{
			this.cmbNomina.DropDownStyle = UltraComboStyle.DropDown;
			this.cmbNomina.Value = System.DBNull.Value;

			this.cmbNomina.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("Exec APListaRealiza {0}, {1}", MenuLatinium.IdUsuario, (int)this.cmbResponsable.Value));
			
			if (this.cmbNomina.Rows.Count == 1)
			{
				this.cmbNomina.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec APRetornaIdRealiza {0}, {1}", MenuLatinium.IdUsuario, (int)this.cmbResponsable.Value));
				this.cmbNomina.DropDownStyle = UltraComboStyle.DropDownList;
			}
		}

		private void cmbResponsable_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
