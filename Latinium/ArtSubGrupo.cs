using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ArtSubGrupo.
	/// </summary>
	public class ArtSubGrupo : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsSubGrupo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPolitica;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoPoliticaObtenerTodos;

		private C1DataRow drSeteoF;

		public ArtSubGrupo()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSubGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IncrPI");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IncrPA");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcEntrada");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcMinEntrada");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pulgadas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPreciador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IncrPVP");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GarantiaExtendida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeGarantiaExtendida");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoGarantiaExtendida");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ServicioDomiciio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ActasRecepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NuevoRecuperacionRemates");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaV");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObProduccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CobroPrimeraCuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescTCSI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Funda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProductosPorLocal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr9");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr12");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr15");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr18");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr21");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr24");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPolitica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroPolitica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoPolitica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroPolitica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mensaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			this.cdsSubGrupo = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbPolitica = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.udsRequerimientoPoliticaObtenerTodos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cdsSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPolitica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPoliticaObtenerTodos)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsSubGrupo
			// 
			this.cdsSubGrupo.BindingContextCtrl = this;
			this.cdsSubGrupo.DataLibrary = "LibFacturacion";
			this.cdsSubGrupo.DataLibraryUrl = "";
			this.cdsSubGrupo.DataSetDef = "dsArticuloTabla";
			this.cdsSubGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSubGrupo.Name = "cdsSubGrupo";
			this.cdsSubGrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSubGrupo.SchemaDef = null;
			this.cdsSubGrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSubGrupo_BeforeFill);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "ArticuloSubGrupo";
			this.ultraGrid1.DataSource = this.cdsSubGrupo;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 200;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.Header.VisiblePosition = 10;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 208;
			ultraGridColumn4.Header.VisiblePosition = 11;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 129;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 101;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Incremento PI";
			ultraGridColumn6.Header.VisiblePosition = 12;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 103;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Incremento PA";
			ultraGridColumn7.Header.VisiblePosition = 13;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 106;
			ultraGridColumn8.Header.Caption = "Ent";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 70;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance4;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "% Ent";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 70;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "% Min. Ent.";
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 70;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Width = 19;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Hidden = true;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance6;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Incr PVP";
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Width = 100;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.Caption = "GEx";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 70;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance7;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "% GEx";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 70;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance8;
			ultraGridColumn18.Header.Caption = "Tiempo GEx";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 70;
			ultraGridColumn19.Header.Caption = "Servicio a Domiciio";
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 70;
			ultraGridColumn20.Header.Caption = "Actas de Recepcion";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 70;
			ultraGridColumn21.Header.Caption = "Activo";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 70;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 20;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 56;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 17;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 23;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 23;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 30;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 20;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 40;
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 28;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 28;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Width = 20;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 38;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance9;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.Width = 100;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance10;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn35.Width = 100;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance11;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.VisiblePosition = 35;
			ultraGridColumn36.Width = 100;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance12;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 36;
			ultraGridColumn37.Width = 100;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance13;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.VisiblePosition = 37;
			ultraGridColumn38.Width = 100;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance14;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 38;
			ultraGridColumn39.Width = 100;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance15;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.VisiblePosition = 39;
			ultraGridColumn40.Width = 100;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance16;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Width = 100;
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
																										 ultraGridColumn41});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8.5F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(0, 0);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1087, 256);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(115, 576);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(87, 29);
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(10, 576);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(96, 29);
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
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
			// cmbPolitica
			// 
			this.cmbPolitica.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn42.Header.VisiblePosition = 0;
			ultraGridColumn43.Header.VisiblePosition = 1;
			ultraGridColumn44.Header.VisiblePosition = 2;
			ultraGridColumn45.Header.VisiblePosition = 4;
			ultraGridColumn46.Header.VisiblePosition = 5;
			ultraGridColumn47.Header.VisiblePosition = 3;
			ultraGridColumn48.Header.VisiblePosition = 6;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48});
			this.cmbPolitica.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPolitica.DisplayMember = "Descripcion";
			this.cmbPolitica.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPolitica.Location = new System.Drawing.Point(336, 104);
			this.cmbPolitica.Name = "cmbPolitica";
			this.cmbPolitica.Size = new System.Drawing.Size(67, 83);
			this.cmbPolitica.TabIndex = 1064;
			this.cmbPolitica.ValueMember = "idRequerimientoPolitica";
			this.cmbPolitica.Visible = false;
			// 
			// udsRequerimientoPoliticaObtenerTodos
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn7.DataType = typeof(int);
			this.udsRequerimientoPoliticaObtenerTodos.Band.Columns.AddRange(new object[] {
																																										 ultraDataColumn1,
																																										 ultraDataColumn2,
																																										 ultraDataColumn3,
																																										 ultraDataColumn4,
																																										 ultraDataColumn5,
																																										 ultraDataColumn6,
																																										 ultraDataColumn7});
			// 
			// ArtSubGrupo
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(1106, 610);
			this.ControlBox = false;
			this.Controls.Add(this.cmbPolitica);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ArtSubGrupo";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "SubGrupos de Artículos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ArtSubGrupo_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ArtSubGrupo_Closing);
			this.Load += new System.EventHandler(this.ArtSubGrupo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPolitica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPoliticaObtenerTodos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSubGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSubGrupo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					if (Convert.ToDecimal(dr.Cells["IncrPVP"].Value) == 0.00m)
					{
						MessageBox.Show(string.Format("Escriba el porcentaje de incremento para Crédito del subgrupo {0}.", dr.Cells["Subgrupo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveCell = dr.Cells[5];
						return;
					}					
				}

				this.cdsSubGrupo.Update();

				if (DialogResult.Yes == MessageBox.Show("¿Desea Actualizar el Precio de Credito Directo?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					Funcion.EjecutaSQL(cdsSeteoF, "Exec ActualizaPreciosCreditoDesdeSubgrupo");
				}
				
				/*string sSQL = "";
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					int idSubgrupo = (int)dr.Cells["idSubGrupo"].Value;
					sSQL = string.Format("Exec ActualizaIncrementoPrecio {0}", idSubgrupo);
					Funcion.EjecutaSQL(cdsSubGrupo, sSQL, true);
				}*/

				this.DialogResult = DialogResult.OK;
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al grabar");
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void VerificaBorrar(string IdArticulo, string stArticulo)
		{
			// Verificacion de Existencia en compras
			string stSelect = "SELECT COUNT(*) FROM Articulo WHERE Articulo.idSubGrupo = "
				+ IdArticulo;
			this.cdsSubGrupo.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = this.cdsSubGrupo.Schema.Connections[0];
			SqlCommand cmdCodigo = new SqlCommand (
				stSelect, (SqlConnection) miConeccion.DbConnection);
			int iCuenta = (int) cmdCodigo.ExecuteScalar();
			if (iCuenta > 0) 
			{
				MessageBox.Show("Tiene " + iCuenta.ToString().Trim() + " Artículos con el Arancel "
					+ stArticulo + ". Eliminación CANCELADA",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cdsSubGrupo.Schema.Connections[0].Close();
				return;
			}
			this.cdsSubGrupo.Schema.Connections[0].Close();
			C1.Data.C1DataRow dr = this.cdsSubGrupo.TableViews["ArticuloSubGrupo"].Rows.Find(int.Parse(IdArticulo));
			dr.Delete();
			// Fin de verificacion de Existencia de cliente en compras
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg = false;
			e.Cancel = true;
			string stCiudad = " SubGrupos?";
			if (e.Rows.Length == 1) stCiudad = " SubGrupo?";
			if(MessageBox.Show("¿Desea Borrar " + e.Rows.Length.ToString().Trim() 
				+ stCiudad, "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
			{
				return;
			}

			Infragistics.Win.UltraWinGrid.UltraGridRow[] drCiudad = e.Rows;
			string ICiudad;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in drCiudad)
			{
				ICiudad = dr.Cells["idSubGrupo"].Text.Trim();
				stCiudad = dr.Cells["SubGrupo"].Text.Trim();
				this.VerificaBorrar(ICiudad, stCiudad);
			}
		}

		private void ArtSubGrupo_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			string stAudita = "Exec AuditaCrear 53, 6, 'Art. subGrupo'";
			Funcion.EjecutaSQL(cdsSubGrupo, stAudita, true);
		}

		private void ArtSubGrupo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Art. SubGrupo'";
			Funcion.EjecutaSQL(cdsSubGrupo, stAudita, true);
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			
		}

		private void ArtSubGrupo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			switch(e.KeyValue)
			{				
				#region Page Down Abre lista desplegable
				case (int) Keys.PageDown:
					if(ultraGrid1.ActiveCell.DroppedDown == false)
						ultraGrid1.ActiveCell.DroppedDown = true;
					break;
				#endregion Page Down Abre lista desplegable
				#region Flechas
				case (int) Keys.Up:
					if (ultraGrid1.ActiveCell.DroppedDown) return;

					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Left:
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (ultraGrid1.ActiveCell.DroppedDown) return;
			
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				#endregion Flechas
				#region Enter
				case (int) Keys.Enter:
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
					//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
				break;
				#endregion Enter
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			
			#region Decimales
			stFormato = stMoneda + stCeros.Substring(0, 2);
			e.Layout.Bands[0].Columns["IncrPI"].Format = stFormato;
			e.Layout.Bands[0].Columns["IncrPA"].Format = stFormato;
			e.Layout.Bands[0].Columns["IncrPVP"].Format = stFormato;
			e.Layout.Bands[0].Columns["PorcEntrada"].Format = stFormato;
			e.Layout.Bands[0].Columns["PorcMinEntrada"].Format = stFormato;
			e.Layout.Bands[0].Columns["PorcentajeGarantiaExtendida"].Format = stFormato;

			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["IncrPI"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["IncrPA"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["IncrPVP"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["PorcEntrada"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["PorcMinEntrada"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["PorcentajeGarantiaExtendida"].MaskInput = stInput;
			#endregion Decimales
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
	}
}
