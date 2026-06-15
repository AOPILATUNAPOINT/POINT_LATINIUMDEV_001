using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Diagnostics;
using Infragistics.Win.UltraWinGrid;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCatalogoHabitaciones.
	/// </summary>
	public class frmCatalogoHabitaciones : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempo;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTipoHabitacion;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcReportes;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl TbBasico;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugdDocumento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsDocumento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSuit;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTubo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRemodelada;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSillaM;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkYacuzzi;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkChimenea;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSillaT;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCocoon;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTurco;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDiseño;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCatalogohabitaciones;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecio;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCatalogoHabitaciones()
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
		string txtDirectorioarchivo = "";

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdImagenHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdCatalogohabitaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlCompartido", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDocumento", 0);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCatalogoHabitaciones));
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEliminar", 1);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 2);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdImagenHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCatalogohabitaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlCompartido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlNube");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			this.TbBasico = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.chkSuit = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkTubo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkRemodelada = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSillaM = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkYacuzzi = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkChimenea = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSillaT = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCocoon = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkTurco = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDiseño = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ugdDocumento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsDocumento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtTipoHabitacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.txtTiempo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.utcReportes = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.txtIdCatalogohabitaciones = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPrecio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.TbBasico.SuspendLayout();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTipoHabitacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).BeginInit();
			this.utcReportes.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCatalogohabitaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			this.SuspendLayout();
			// 
			// TbBasico
			// 
			this.TbBasico.Controls.Add(this.chkSuit);
			this.TbBasico.Controls.Add(this.chkTubo);
			this.TbBasico.Controls.Add(this.chkRemodelada);
			this.TbBasico.Controls.Add(this.chkSillaM);
			this.TbBasico.Controls.Add(this.chkYacuzzi);
			this.TbBasico.Controls.Add(this.chkChimenea);
			this.TbBasico.Controls.Add(this.chkSillaT);
			this.TbBasico.Controls.Add(this.chkCocoon);
			this.TbBasico.Controls.Add(this.chkTurco);
			this.TbBasico.Controls.Add(this.chkDiseño);
			this.TbBasico.Location = new System.Drawing.Point(-10000, -10000);
			this.TbBasico.Name = "TbBasico";
			this.TbBasico.Size = new System.Drawing.Size(516, 206);
			// 
			// chkSuit
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSuit.Appearance = appearance1;
			this.chkSuit.BackColor = System.Drawing.Color.Transparent;
			this.chkSuit.Enabled = false;
			this.chkSuit.Location = new System.Drawing.Point(24, 24);
			this.chkSuit.Name = "chkSuit";
			this.chkSuit.Size = new System.Drawing.Size(104, 22);
			this.chkSuit.TabIndex = 1022;
			this.chkSuit.Text = "Suit Loft 2 Pisos";
			// 
			// chkTubo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTubo.Appearance = appearance2;
			this.chkTubo.BackColor = System.Drawing.Color.Transparent;
			this.chkTubo.Enabled = false;
			this.chkTubo.Location = new System.Drawing.Point(24, 80);
			this.chkTubo.Name = "chkTubo";
			this.chkTubo.Size = new System.Drawing.Size(56, 22);
			this.chkTubo.TabIndex = 1021;
			this.chkTubo.Text = "Tubo";
			// 
			// chkRemodelada
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkRemodelada.Appearance = appearance3;
			this.chkRemodelada.BackColor = System.Drawing.Color.Transparent;
			this.chkRemodelada.Enabled = false;
			this.chkRemodelada.Location = new System.Drawing.Point(160, 144);
			this.chkRemodelada.Name = "chkRemodelada";
			this.chkRemodelada.Size = new System.Drawing.Size(112, 22);
			this.chkRemodelada.TabIndex = 1020;
			this.chkRemodelada.Text = "Remodelada";
			// 
			// chkSillaM
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSillaM.Appearance = appearance4;
			this.chkSillaM.BackColor = System.Drawing.Color.Transparent;
			this.chkSillaM.Enabled = false;
			this.chkSillaM.Location = new System.Drawing.Point(376, 80);
			this.chkSillaM.Name = "chkSillaM";
			this.chkSillaM.Size = new System.Drawing.Size(96, 22);
			this.chkSillaM.TabIndex = 1019;
			this.chkSillaM.Text = "Silla Madam T";
			// 
			// chkYacuzzi
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkYacuzzi.Appearance = appearance5;
			this.chkYacuzzi.BackColor = System.Drawing.Color.Transparent;
			this.chkYacuzzi.Enabled = false;
			this.chkYacuzzi.Location = new System.Drawing.Point(272, 24);
			this.chkYacuzzi.Name = "chkYacuzzi";
			this.chkYacuzzi.Size = new System.Drawing.Size(72, 22);
			this.chkYacuzzi.TabIndex = 1018;
			this.chkYacuzzi.Text = "Yacuzzi";
			// 
			// chkChimenea
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkChimenea.Appearance = appearance6;
			this.chkChimenea.BackColor = System.Drawing.Color.Transparent;
			this.chkChimenea.Enabled = false;
			this.chkChimenea.Location = new System.Drawing.Point(24, 144);
			this.chkChimenea.Name = "chkChimenea";
			this.chkChimenea.Size = new System.Drawing.Size(96, 22);
			this.chkChimenea.TabIndex = 1017;
			this.chkChimenea.Text = "Chimenea";
			// 
			// chkSillaT
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSillaT.Appearance = appearance7;
			this.chkSillaT.BackColor = System.Drawing.Color.Transparent;
			this.chkSillaT.Enabled = false;
			this.chkSillaT.Location = new System.Drawing.Point(272, 80);
			this.chkSillaT.Name = "chkSillaT";
			this.chkSillaT.Size = new System.Drawing.Size(80, 22);
			this.chkSillaT.TabIndex = 1016;
			this.chkSillaT.Text = "Silla Tantra";
			// 
			// chkCocoon
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCocoon.Appearance = appearance8;
			this.chkCocoon.BackColor = System.Drawing.Color.Transparent;
			this.chkCocoon.Enabled = false;
			this.chkCocoon.Location = new System.Drawing.Point(160, 80);
			this.chkCocoon.Name = "chkCocoon";
			this.chkCocoon.Size = new System.Drawing.Size(96, 22);
			this.chkCocoon.TabIndex = 1015;
			this.chkCocoon.Text = "Cocoon Cama";
			// 
			// chkTurco
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTurco.Appearance = appearance9;
			this.chkTurco.BackColor = System.Drawing.Color.Transparent;
			this.chkTurco.Enabled = false;
			this.chkTurco.Location = new System.Drawing.Point(376, 24);
			this.chkTurco.Name = "chkTurco";
			this.chkTurco.Size = new System.Drawing.Size(56, 22);
			this.chkTurco.TabIndex = 1014;
			this.chkTurco.Text = "Turco";
			// 
			// chkDiseño
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDiseño.Appearance = appearance10;
			this.chkDiseño.BackColor = System.Drawing.Color.Transparent;
			this.chkDiseño.Enabled = false;
			this.chkDiseño.Location = new System.Drawing.Point(160, 24);
			this.chkDiseño.Name = "chkDiseño";
			this.chkDiseño.Size = new System.Drawing.Size(88, 22);
			this.chkDiseño.TabIndex = 1013;
			this.chkDiseño.Text = "Diseño Autor";
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.ugdDocumento);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(516, 206);
			// 
			// ugdDocumento
			// 
			this.ugdDocumento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugdDocumento.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugdDocumento.DataSource = this.udsDocumento;
			appearance11.BackColor = System.Drawing.Color.White;
			this.ugdDocumento.DisplayLayout.Appearance = appearance11;
			this.ugdDocumento.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 137;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 222;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Width = 305;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Width = 83;
			appearance12.Image = ((object)(resources.GetObject("appearance12.Image")));
			appearance12.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance12;
			appearance13.Image = ((object)(resources.GetObject("appearance13.Image")));
			ultraGridColumn5.CellButtonAppearance = appearance13;
			ultraGridColumn5.Header.Caption = "...";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn5.Width = 41;
			appearance14.Image = ((object)(resources.GetObject("appearance14.Image")));
			appearance14.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance14;
			appearance15.Image = ((object)(resources.GetObject("appearance15.Image")));
			ultraGridColumn6.CellButtonAppearance = appearance15;
			ultraGridColumn6.Header.Caption = "...";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn6.Width = 28;
			appearance16.Image = ((object)(resources.GetObject("appearance16.Image")));
			appearance16.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance16;
			appearance17.Image = ((object)(resources.GetObject("appearance17.Image")));
			ultraGridColumn7.CellButtonAppearance = appearance17;
			ultraGridColumn7.Header.Caption = "...";
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn7.Width = 34;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.ugdDocumento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ugdDocumento.DisplayLayout.GroupByBox.Hidden = true;
			this.ugdDocumento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugdDocumento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugdDocumento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.ugdDocumento.DisplayLayout.Override.CardAreaAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 10F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugdDocumento.DisplayLayout.Override.HeaderAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugdDocumento.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.ugdDocumento.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugdDocumento.Enabled = false;
			this.ugdDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugdDocumento.Location = new System.Drawing.Point(0, 16);
			this.ugdDocumento.Name = "ugdDocumento";
			this.ugdDocumento.Size = new System.Drawing.Size(512, 184);
			this.ugdDocumento.TabIndex = 811;
			this.ugdDocumento.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugdDocumento_ClickCellButton);
			// 
			// udsDocumento
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			this.udsDocumento.Band.Columns.AddRange(new object[] {
																														 ultraDataColumn1,
																														 ultraDataColumn2,
																														 ultraDataColumn3,
																														 ultraDataColumn4});
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Enabled = false;
			this.lblCliente.Location = new System.Drawing.Point(16, 112);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(63, 16);
			this.lblCliente.TabIndex = 990;
			this.lblCliente.Text = "Descripción";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance22;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(120, 104);
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(336, 72);
			this.txtDescripcion.TabIndex = 989;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(16, 64);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(99, 16);
			this.label1.TabIndex = 992;
			this.label1.Text = "Tipo de Habitacion";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTipoHabitacion
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTipoHabitacion.Appearance = appearance23;
			this.txtTipoHabitacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTipoHabitacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTipoHabitacion.Enabled = false;
			this.txtTipoHabitacion.Location = new System.Drawing.Point(120, 64);
			this.txtTipoHabitacion.Name = "txtTipoHabitacion";
			this.txtTipoHabitacion.Size = new System.Drawing.Size(336, 21);
			this.txtTipoHabitacion.TabIndex = 991;
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnConsultar.BackColor = System.Drawing.SystemColors.Control;
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(88, 512);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 998;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevo.BackColor = System.Drawing.SystemColors.Control;
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(16, 512);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 24);
			this.btnNuevo.TabIndex = 997;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnEditar.BackColor = System.Drawing.SystemColors.Control;
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(168, 512);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 24);
			this.btnEditar.TabIndex = 996;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-448, 408);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1312, 88);
			this.groupBox1.TabIndex = 995;
			this.groupBox1.TabStop = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancelar.BackColor = System.Drawing.SystemColors.Control;
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(328, 512);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 24);
			this.btnCancelar.TabIndex = 994;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.BackColor = System.Drawing.SystemColors.Control;
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(248, 512);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 993;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// txtTiempo
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempo.Appearance = appearance24;
			this.txtTiempo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempo.Enabled = false;
			this.txtTiempo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTiempo.FormatString = "";
			this.txtTiempo.Location = new System.Drawing.Point(120, 200);
			this.txtTiempo.MaxValue = 100;
			this.txtTiempo.MinValue = 0;
			this.txtTiempo.Name = "txtTiempo";
			this.txtTiempo.PromptChar = ' ';
			this.txtTiempo.Size = new System.Drawing.Size(112, 22);
			this.txtTiempo.TabIndex = 1000;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(16, 200);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(41, 17);
			this.label14.TabIndex = 999;
			this.label14.Text = "Tiempo";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(272, 200);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 17);
			this.label2.TabIndex = 1001;
			this.label2.Text = "Precio";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// utcReportes
			// 
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.utcReportes.ActiveTabAppearance = appearance25;
			this.utcReportes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance26.FontData.Name = "Tahoma";
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.White;
			this.utcReportes.Appearance = appearance26;
			this.utcReportes.BackColor = System.Drawing.Color.Gainsboro;
			this.utcReportes.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcReportes.Controls.Add(this.TbBasico);
			this.utcReportes.Controls.Add(this.ultraTabPageControl10);
			this.utcReportes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.utcReportes.Location = new System.Drawing.Point(0, 256);
			this.utcReportes.Name = "utcReportes";
			this.utcReportes.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcReportes.Size = new System.Drawing.Size(520, 232);
			this.utcReportes.TabButtonStyle = Infragistics.Win.UIElementButtonStyle.Flat;
			this.utcReportes.TabIndex = 1060;
			ultraTab1.TabPage = this.TbBasico;
			ultraTab1.Text = "Caracteristicas";
			ultraTab2.TabPage = this.ultraTabPageControl10;
			ultraTab2.Text = "Imagenes";
			this.utcReportes.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													ultraTab1,
																																													ultraTab2});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(516, 206);
			// 
			// txtIdCatalogohabitaciones
			// 
			this.txtIdCatalogohabitaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCatalogohabitaciones.Enabled = false;
			this.txtIdCatalogohabitaciones.Location = new System.Drawing.Point(472, 16);
			this.txtIdCatalogohabitaciones.Name = "txtIdCatalogohabitaciones";
			this.txtIdCatalogohabitaciones.PromptChar = ' ';
			this.txtIdCatalogohabitaciones.Size = new System.Drawing.Size(24, 21);
			this.txtIdCatalogohabitaciones.TabIndex = 1061;
			this.txtIdCatalogohabitaciones.Visible = false;
			// 
			// txtPrecio
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecio.Appearance = appearance27;
			this.txtPrecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecio.Enabled = false;
			this.txtPrecio.FormatString = "#,##0.00";
			this.txtPrecio.Location = new System.Drawing.Point(336, 200);
			this.txtPrecio.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtPrecio.Name = "txtPrecio";
			this.txtPrecio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecio.PromptChar = ' ';
			this.txtPrecio.Size = new System.Drawing.Size(80, 21);
			this.txtPrecio.TabIndex = 1062;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(16, 16);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(43, 16);
			this.lblBodega.TabIndex = 1064;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 30;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 60;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Width = 180;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 60;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(120, 16);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 21);
			this.cmbBodega.TabIndex = 1063;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// frmCatalogoHabitaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(512, 558);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtPrecio);
			this.Controls.Add(this.txtIdCatalogohabitaciones);
			this.Controls.Add(this.utcReportes);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtTiempo);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtTipoHabitacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtDescripcion);
			this.MaximizeBox = false;
			this.Name = "frmCatalogoHabitaciones";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "frmCatalogoHabitaciones";
			this.Load += new System.EventHandler(this.frmCatalogoHabitaciones_Load);
			this.TbBasico.ResumeLayout(false);
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ugdDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTipoHabitacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).EndInit();
			this.utcReportes.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtIdCatalogohabitaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.txtIdCatalogohabitaciones.Value = 0;
			#region HabilitaBotones
			//this.btnCancelar_Click(sender, e);
			this.txtDescripcion.Enabled = true;
			this.txtTipoHabitacion.Enabled = true;
			this.txtTiempo.Enabled = true;
			this.cmbBodega.Enabled = true;
			this.txtPrecio.Enabled = true;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.chkSuit.Enabled = true;
			this.chkDiseño.Enabled = true;
			this.chkYacuzzi.Enabled = true;
			this.chkTubo.Enabled = true;
			this.chkTurco.Enabled = true;
			this.chkCocoon.Enabled = true;
			this.chkSillaT.Enabled = true;
			this.chkSillaM.Enabled = true;
			this.chkChimenea.Enabled = true;
			this.chkRemodelada.Enabled = true;									
			FuncionesProcedimientos.EstadoGrids(true, this.ugdDocumento);
			#endregion HabilitaBotones
		
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			#region DesabilitaBotones 
			this.txtDescripcion.Value = "" ;
			this.txtTipoHabitacion.Value = "";
			this.txtPrecio.Value = 0;
			this.txtTiempo.Value = 0;
			this.cmbBodega.Value = 0;
			this.cmbBodega.Enabled = false;
			this.txtTiempo.Enabled = false;
			this.txtDescripcion.Enabled = false;
			this.txtTipoHabitacion.Enabled = false;
			this.txtPrecio.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnCancelar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnEditar.Enabled = false;

			this.chkSuit.Checked = false;
			this.chkDiseño.Checked = false;
			this.chkYacuzzi.Checked = false;
			this.chkTubo.Checked = false;
			this.chkTurco.Checked = false;
			this.chkCocoon.Checked = false;
			this.chkSillaT.Checked = false;
			this.chkSillaM.Checked = false;
			this.chkChimenea.Checked = false;
			this.chkRemodelada.Checked = false;

			this.chkSuit.Enabled = false;
			this.chkDiseño.Enabled = false;
			this.chkYacuzzi.Enabled = false;
			this.chkTubo.Enabled = false;
			this.chkTurco.Enabled = false;
			this.chkCocoon.Enabled = false;
			this.chkSillaT.Enabled = false;
			this.chkSillaM.Enabled = false;
			this.chkChimenea.Enabled = false;
			this.chkRemodelada.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(false, this.ugdDocumento);

			try
			{
				if (this.ugdDocumento.Rows.Count > 0)
				{
					for (int i = this.ugdDocumento.Rows.Count - 1; i >= 0; i--)
					{
						if (!this.ugdDocumento.Rows[i].IsAddRow)
							this.ugdDocumento.Rows[i].Delete(false);
					}
				}
				this.ugdDocumento.Enabled = false;
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al limpiar las filas del grid: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

			#endregion DesabilitaBotones

		 }
		private void Consulta(int IdCatalogohabitaciones)
		{
			string sSQL = string.Format("Exec ConsultaCompletaCatalogo {0}", IdCatalogohabitaciones);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{	
				this.txtIdCatalogohabitaciones.Value = dr.GetInt32(0);
				this.txtTipoHabitacion.Value = dr.GetString(1);
				this.txtDescripcion.Value = dr.GetString(2);
				this.txtTiempo.Value = dr.GetString(3);
				this.txtPrecio.Value = dr.GetDecimal(4);
				this.chkSuit.Checked = dr.GetBoolean(5);
				this.chkDiseño.Checked = dr.GetBoolean(6);
				this.chkYacuzzi.Checked = dr.GetBoolean(7);
				this.chkTurco.Checked = dr.GetBoolean(8);
				this.chkTubo.Checked = dr.GetBoolean(9);
				this.chkCocoon.Checked = dr.GetBoolean(10);
				this.chkSillaT.Checked = dr.GetBoolean(11);
				this.chkSillaM.Checked = dr.GetBoolean(12);
				this.chkChimenea.Checked = dr.GetBoolean(13);
				this.chkRemodelada.Checked = dr.GetBoolean(14);
				this.cmbBodega.Value = dr.GetInt32(15);
			}
			dr.Close();

			string sSQLImg = string.Format("Exec ConsultaImagenesHabitaciones {0}",this.txtIdCatalogohabitaciones.Value);
			
			this.ugdDocumento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLImg);
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			using (frmBsucaCatalogoHabitaciones Busqueda = new frmBsucaCatalogoHabitaciones())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.grdFacturas.ActiveRow.Cells["IdCatalogohabitaciones"].Value);
				}
			}

			this.btnEditar.Enabled = true;
			this.btnNuevo.Enabled = true;
			this.btnCancelar.Enabled = false;
			this.btnGuardar.Enabled = false;
		}

		private void frmCatalogoHabitaciones_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "1131");
			
			if (!Funcion.Permiso("2154", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar al Catalogo de Habitaciones", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;

		}

		private void ultraCheckEditor7_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			#region HabilitaBotones
			//this.btnCancelar_Click(sender, e);
			this.txtDescripcion.Enabled = true;
			this.txtTipoHabitacion.Enabled = true;
			this.txtTiempo.Enabled = true;
			this.txtPrecio.Enabled = true;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnNuevo.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.chkSuit.Enabled = true;
			this.chkDiseño.Enabled = true;
			this.chkYacuzzi.Enabled = true;
			this.chkTubo.Enabled = true;
			this.chkTurco.Enabled = true;
			this.chkCocoon.Enabled = true;
			this.chkSillaT.Enabled = true;
			this.chkSillaM.Enabled = true;
			this.chkChimenea.Enabled = true;
			this.chkRemodelada.Enabled = true;	
			this.cmbBodega.Enabled = true;					
			FuncionesProcedimientos.EstadoGrids(true, this.ugdDocumento);
			#endregion HabilitaBotones
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;

			if (this.txtTipoHabitacion.Text == null || this.txtTipoHabitacion.Text.Trim() == "")
			{
				MessageBox.Show("Ingrese el Tipo de Habitacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTipoHabitacion.Focus();
				return;
			}

			if (this.txtDescripcion.Text == null || this.txtDescripcion.Text.Trim() == "")
			{
				MessageBox.Show("Ingrese la descripcion de la Habitacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDescripcion.Focus();
				return;
			}

			int tiempo = 0;
			try
			{
				tiempo = Convert.ToInt32(this.txtTiempo.Text);
			}
			catch
			{
				MessageBox.Show("El tiempo debe ser un número válido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTiempo.Focus();
				return;
			}

			if (tiempo <= 0)
			{
				MessageBox.Show("El tiempo debe ser mayor a CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTiempo.Focus();
				return;
			}

			decimal precio = 0;
			try
			{
				precio = Convert.ToDecimal(this.txtPrecio.Text);
			}
			catch
			{
				MessageBox.Show("El precio debe ser un número decimal válido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtPrecio.Focus();
				return;
			}

			if (precio <= 0)
			{
				MessageBox.Show("El Precio debe ser mayor a CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtPrecio.Focus();
				return;
			}
			#endregion

			// Guardar la habitación
			string sSQL = string.Format("Exec InsertaCatalogoHabitaciones {0}, '{1}', '{2}' , {3} , {4} , {5} , {6} , {7} , {8} , {9} , {10} , {11} , {12} , '{13}' , '{14}',{15}",
				(int)this.txtIdCatalogohabitaciones.Value,
				this.txtTipoHabitacion.Text.ToString(),
				this.txtDescripcion.Text.ToString(),
				this.chkSuit.Checked,
				this.chkDiseño.Checked,
				this.chkYacuzzi.Checked,
				this.chkTurco.Checked,
				this.chkTubo.Checked,
				this.chkCocoon.Checked,
				this.chkSillaT.Checked,
				this.chkSillaM.Checked,
				this.chkChimenea.Checked,
				this.chkRemodelada.Checked,
				this.txtTiempo.Text.ToString(),
				this.txtPrecio.Text.ToString(),
				(int)this.cmbBodega.Value);

			this.txtIdCatalogohabitaciones.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			string tipoHabitacion = this.txtTipoHabitacion.Text.Trim().Replace(" ", "_");
			string basePath = "\\\\192.168.1.10\\Proyectos\\ImagenesHabitaciones";
			string carpetaCatalogo = basePath + "\\" + tipoHabitacion;



			try
			{
				Process.Start("cmd.exe", "/C net use \\\\192.168.1.10 /user:latinium 10dgpv15.p$");
			}
			catch {}

			if (!Directory.Exists(carpetaCatalogo))
			{
				try
				{
					Directory.CreateDirectory(carpetaCatalogo);
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error al crear carpeta: " + ex.Message);
					return;
				}
			}

			ugdDocumento.PerformAction(UltraGridAction.ExitEditMode);
			ugdDocumento.UpdateData(); 

			foreach (UltraGridRow row in ugdDocumento.Rows)
			{
				if (!row.IsAddRow)
				{
					row.Update();
				}
			}
			foreach (UltraGridRow dr in ugdDocumento.Rows.All)
			{
				if (dr.IsAddRow || dr.Cells["UrlCompartido"].Value == null || dr.Cells["UrlCompartido"].Value.ToString().Trim() == "")
					continue;

				string archivoOrigen = dr.Cells["UrlCompartido"].Value.ToString();

				if (!File.Exists(archivoOrigen))
				{
					MessageBox.Show("Archivo no encontrado: " + archivoOrigen, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					continue;
				}

				string nombreArchivo = Path.GetFileName(archivoOrigen).Replace(" ", "_");
				string destino = Path.Combine(carpetaCatalogo, nombreArchivo);
				string rutaRelativa = tipoHabitacion + "\\" + nombreArchivo;

				if (!File.Exists(destino))
				{
					File.Copy(archivoOrigen, destino, false);
				}

				dr.Cells["UrlNube"].Value = rutaRelativa;

				string destinoSQL = destino.Replace("'", "''");
				string rutaRelativaSQL = rutaRelativa.Replace("'", "''");

				string sSQLImagenes = string.Format(
					"EXEC [grabaUrlNubeCatalogoHabitaciones] {0}, '{1}', '{2}'",
					(int)this.txtIdCatalogohabitaciones.Value,
					destinoSQL,
					rutaRelativaSQL);

				Funcion.EjecutaSQL(cdsSeteoF, sSQLImagenes);
			}

			MessageBox.Show("Datos de la Habitación guardados correctamente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			this.btnCancelar_Click(sender, e);
		}


		private void txtPrecio_Enter(object sender, System.EventArgs e)
		{


		}

		private void txtPrecio_Leave(object sender, System.EventArgs e)
		{

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtPrecio_Leave_1(object sender, System.EventArgs e)
		{
		
		}

		private void txtPrecio_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private string[] imagenesSeleccionadas = null;
		private void ugdDocumento_ClickCellButton(object sender, CellEventArgs e)
		{
			#region imagenes
			if (e.Cell.Column.ToString() == "btnDocumento")
			{
				try
				{
					OpenFileDialog buscar = new OpenFileDialog();
					buscar.Filter = "Imágenes (*.png;*.jpg)|*.png;*.jpg";
					buscar.Multiselect = true;

					if (buscar.ShowDialog() == DialogResult.OK)
					{
						int agregadas = 0;

						ugdDocumento.PerformAction(UltraGridAction.ExitEditMode);
						ugdDocumento.UpdateData();

					#region Seleccion Imagen
						foreach (string archivo in buscar.FileNames)
						{
							if (archivo == null || archivo.Trim() == "" || !File.Exists(archivo))
								continue;

							string rutaArchivo = archivo;
							string nombreArchivo = Path.GetFileName(rutaArchivo).Replace(" ", "_");

							bool yaExiste = false;

							foreach (UltraGridRow row in ugdDocumento.Rows)
							{
								if (row.Cells["UrlCompartido"].Value != null &&
									row.Cells["UrlCompartido"].Value.ToString().ToLower() == rutaArchivo.ToLower())
								{
									yaExiste = true;
									break;
								}
							}
							#region Validacion Imagen

							if (!yaExiste)
							{
								UltraGridRow dr = ugdDocumento.DisplayLayout.Bands[0].AddNew();
								dr.Cells["UrlCompartido"].Value = rutaArchivo;
								dr.Cells["UrlNube"].Value = "";

								ugdDocumento.PerformAction(UltraGridAction.ExitEditMode);
								dr.Update();
								agregadas++;
							}
							#endregion Validacion Imagen
						}

						ugdDocumento.UpdateData();
				 #endregion Seleccion Imagen
						if (agregadas > 0)
						{
							MessageBox.Show("Imágenes seleccionadas. Presione 'Guardar' para almacenarlas.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
						else
						{
							MessageBox.Show("Las imágenes ya estaban agregadas.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
						}
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error al seleccionar imágenes: " + ex.Message);
				}
			}
			#endregion imganes

			#region Ver
			if (e.Cell.Column.ToString() == "btnVer")
			{
				try
				{
					string ruta = e.Cell.Row.Cells["UrlCompartido"].Value.ToString();

					if (ruta != null && ruta.Trim().Length > 10 && File.Exists(ruta))
					{
						ProcessStartInfo psi = new ProcessStartInfo();
						psi.FileName = ruta;
						psi.UseShellExecute = true; // Importante para abrir imágenes según la aplicación predeterminada

						Process.Start(psi);
					}
					else
					{
						MessageBox.Show("No se encontró el archivo: " + ruta, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error al abrir imagen: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			#endregion Ver

			#region Eliminar
			if (e.Cell.Column.ToString() == "btnEliminar")
			{
				try
				{
					string ruta = "";
					if (e.Cell.Row.Cells["UrlCompartido"].Value != null)
						ruta = e.Cell.Row.Cells["UrlCompartido"].Value.ToString();
					if (e.Cell.Row.Cells["IdImagenHabitacion"].Value != null &&
						e.Cell.Row.Cells["IdImagenHabitacion"].Value.ToString() != "")
					{

						if (ruta != "" && System.IO.File.Exists(ruta))
						{
							System.IO.File.Delete(ruta);
						}

						int id = Convert.ToInt32(e.Cell.Row.Cells["IdImagenHabitacion"].Value);
						string sSQL = "DELETE FROM ImagenesHabitaciones WHERE IdImagenHabitacion = " + id.ToString();
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);
					}

					e.Cell.Row.Delete(false); 
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error al eliminar imagen: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			#endregion Eliminar
		}
	}
}
