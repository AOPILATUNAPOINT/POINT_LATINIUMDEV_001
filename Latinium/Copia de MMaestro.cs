using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de MMaestro.
	/// </summary>
	public class MMaestro : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsMarceloMaestro;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdTalla;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.Misc.UltraButton btReparte;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMarceloCodigo;
		private DevExpress.XtraEditors.DateEdit dtFecha;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MMaestro()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_MarceloCodigo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloMaestro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla16");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla17");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla19");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla20");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla21");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla22");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla23");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla24");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla25");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla16");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla17");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla19");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla20");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla21");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla22");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla23");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla24");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla25");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unir");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cantidad", 6, true, "_MarceloCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cantidad", 6, true);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.cdsMarceloMaestro = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.grdTalla = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label3 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtMarceloCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btReparte = new Infragistics.Win.Misc.UltraButton();
			this.dtFecha = new DevExpress.XtraEditors.DateEdit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloMaestro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdTalla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMarceloCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsMarceloMaestro
			// 
			this.cdsMarceloMaestro.BindingContextCtrl = this;
			this.cdsMarceloMaestro.DataLibrary = "LibFacturacion";
			this.cdsMarceloMaestro.DataLibraryUrl = "";
			this.cdsMarceloMaestro.DataSetDef = "MarceloMaestroC";
			this.cdsMarceloMaestro.FillOnRequest = false;
			this.cdsMarceloMaestro.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMarceloMaestro.Name = "cdsMarceloMaestro";
			this.cdsMarceloMaestro.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMarceloMaestro.SchemaDef = null;
			this.cdsMarceloMaestro.BeforeFill += new C1.Data.FillEventHandler(this.cdsMarceloMaestro_BeforeFill);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_MarceloMaestro";
			this.barraDato1.DataNombreId = "idMarceloMaestro";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsMarceloMaestro;
			this.barraDato1.DataTabla = "MarceloMaestro";
			this.barraDato1.DataTablaHija = "MarceloCodigo";
			this.barraDato1.Location = new System.Drawing.Point(168, 312);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 138;
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
			// grdTalla
			// 
			this.grdTalla.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdTalla.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdTalla.DataMember = "_MarceloMaestro.MarceloMaestro - MarceloCodigo";
			this.grdTalla.DataSource = this.cdsMarceloMaestro;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdTalla.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "Articulo_Filtro";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 33;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 146;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 167;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 94;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn18.Header.VisiblePosition = 18;
			ultraGridColumn19.Header.VisiblePosition = 19;
			ultraGridColumn20.Header.VisiblePosition = 20;
			ultraGridColumn21.Header.VisiblePosition = 21;
			ultraGridColumn22.Header.VisiblePosition = 22;
			ultraGridColumn23.Header.VisiblePosition = 23;
			ultraGridColumn24.Header.VisiblePosition = 24;
			ultraGridColumn25.Header.VisiblePosition = 25;
			ultraGridColumn26.Header.VisiblePosition = 26;
			ultraGridColumn27.Header.VisiblePosition = 27;
			ultraGridColumn28.Header.VisiblePosition = 28;
			ultraGridColumn29.Header.VisiblePosition = 29;
			ultraGridColumn30.Header.VisiblePosition = 30;
			ultraGridColumn31.Header.VisiblePosition = 31;
			ultraGridColumn32.Header.VisiblePosition = 32;
			ultraGridColumn33.Header.VisiblePosition = 33;
			ultraGridColumn34.Header.VisiblePosition = 34;
			ultraGridColumn35.Header.VisiblePosition = 35;
			ultraGridColumn36.Header.VisiblePosition = 36;
			ultraGridColumn37.Header.VisiblePosition = 37;
			ultraGridColumn38.Header.VisiblePosition = 38;
			ultraGridColumn39.Header.VisiblePosition = 39;
			ultraGridColumn40.Header.VisiblePosition = 40;
			ultraGridColumn41.Header.VisiblePosition = 41;
			ultraGridColumn42.Header.VisiblePosition = 42;
			ultraGridColumn43.Header.VisiblePosition = 43;
			ultraGridColumn44.Header.VisiblePosition = 44;
			ultraGridColumn45.Header.VisiblePosition = 45;
			ultraGridColumn46.Header.VisiblePosition = 46;
			ultraGridColumn47.Header.VisiblePosition = 47;
			ultraGridColumn48.Header.VisiblePosition = 48;
			ultraGridColumn49.Header.VisiblePosition = 49;
			ultraGridColumn50.Header.VisiblePosition = 50;
			ultraGridColumn51.Header.VisiblePosition = 51;
			ultraGridColumn52.Header.VisiblePosition = 52;
			ultraGridColumn53.Header.VisiblePosition = 53;
			ultraGridColumn54.Header.VisiblePosition = 54;
			ultraGridColumn55.Header.VisiblePosition = 55;
			ultraGridColumn56.Header.VisiblePosition = 56;
			ultraGridColumn57.Header.VisiblePosition = 57;
			ultraGridColumn58.Header.VisiblePosition = 2;
			ultraGridColumn59.Header.VisiblePosition = 58;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 59;
			ultraGridColumn60.Hidden = true;
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
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60});
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			this.grdTalla.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdTalla.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdTalla.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdTalla.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grdTalla.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdTalla.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdTalla.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdTalla.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.grdTalla.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdTalla.Location = new System.Drawing.Point(20, 72);
			this.grdTalla.Name = "grdTalla";
			this.grdTalla.Size = new System.Drawing.Size(648, 224);
			this.grdTalla.TabIndex = 137;
			this.grdTalla.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdTalla_KeyDown);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(67, 16);
			this.label3.TabIndex = 140;
			this.label3.Text = "Descripcion:";
			// 
			// txtDescripcion
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance6;
			this.txtDescripcion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarceloMaestro, "_MarceloMaestro.Descripcion"));
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(104, 8);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(536, 21);
			this.txtDescripcion.TabIndex = 139;
			// 
			// txtMarceloCodigo
			// 
			this.txtMarceloCodigo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMarceloCodigo.Appearance = appearance7;
			this.txtMarceloCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarceloMaestro, "_MarceloMaestro.MarceloMaestro - MarceloCodigo.idMarceloCodigo"));
			this.txtMarceloCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMarceloCodigo.Location = new System.Drawing.Point(504, 312);
			this.txtMarceloCodigo.Name = "txtMarceloCodigo";
			this.txtMarceloCodigo.Size = new System.Drawing.Size(44, 21);
			this.txtMarceloCodigo.TabIndex = 141;
			// 
			// btReparte
			// 
			this.btReparte.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btReparte.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btReparte.Location = new System.Drawing.Point(24, 312);
			this.btReparte.Name = "btReparte";
			this.btReparte.Size = new System.Drawing.Size(56, 23);
			this.btReparte.TabIndex = 142;
			this.btReparte.Text = "Repartir";
			this.btReparte.Click += new System.EventHandler(this.btReparte_Click);
			// 
			// dtFecha
			// 
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsMarceloMaestro, "_MarceloMaestro.Fecha"));
			this.dtFecha.EditValue = new System.DateTime(2004, 9, 20, 0, 0, 0, 0);
			this.dtFecha.Location = new System.Drawing.Point(104, 32);
			this.dtFecha.Name = "dtFecha";
			// 
			// dtFecha.Properties
			// 
			this.dtFecha.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFecha.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFecha.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFecha.Properties.EditFormat.FormatString = "dd/MM/yyyy";
			this.dtFecha.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.Custom;
			this.dtFecha.Size = new System.Drawing.Size(96, 21);
			this.dtFecha.TabIndex = 143;
			// 
			// MMaestro
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(688, 349);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btReparte);
			this.Controls.Add(this.txtMarceloCodigo);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.grdTalla);
			this.Name = "MMaestro";
			this.Text = "MMaestro";
			this.Load += new System.EventHandler(this.MMaestro_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloMaestro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdTalla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMarceloCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			this.barraDato1.BorraRegistro();  	
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro(); 
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();  
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();  
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			int t=0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
			{				
				if (dr.Cells["Talla1"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla1"].Value;
				if (dr.Cells["Talla2"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla2"].Value;
				if (dr.Cells["Talla3"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla3"].Value;
				if (dr.Cells["Talla4"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla4"].Value;
				if (dr.Cells["Talla5"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla5"].Value;
				if (dr.Cells["Talla6"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla6"].Value;
				if (dr.Cells["Talla7"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla7"].Value;
				if (dr.Cells["Talla8"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla8"].Value;
				if (dr.Cells["Talla9"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla9"].Value;
				if (dr.Cells["Talla10"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla10"].Value;
				if (dr.Cells["Talla11"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla11"].Value;
				if (dr.Cells["Talla12"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla12"].Value;
				if (dr.Cells["Talla13"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla13"].Value;
				if (dr.Cells["Talla14"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla14"].Value;
				if (dr.Cells["Talla15"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla15"].Value;
				if (dr.Cells["Talla16"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla16"].Value;
				if (dr.Cells["Talla17"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla17"].Value;
				if (dr.Cells["Talla18"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla18"].Value;
				if (dr.Cells["Talla19"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla19"].Value;
				if (dr.Cells["Talla20"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla20"].Value;
				if (dr.Cells["Talla21"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla21"].Value;
				if (dr.Cells["Talla22"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla22"].Value;
				if (dr.Cells["Talla23"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla23"].Value;
				if (dr.Cells["Talla24"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla24"].Value;
				if (dr.Cells["Talla25"].Value != System.DBNull.Value)
					t=t +(int)dr.Cells["Talla25"].Value;

				dr.Cells["Cantidad"].Value =t;
				t=0; 
			}
			this.barraDato1.GrabaRegistro(); 
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
		
		}
			C1.Data.C1DataRow drMarceloMaestro;
		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drMarceloMaestro = cdsMarceloMaestro.TableViews["MarceloMaestro"].Rows[0];
			}
			catch{};
		}

		private void cdsMarceloMaestro_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMarceloMaestro.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void MMaestro_Load(object sender, System.EventArgs e)
		{
			this.barraDato1.HabilitaControles(false);
			this.barraDato1.PosUltima();  
		}

		private void btReparte_Click(object sender, System.EventArgs e)
		{
			MarceloReparteC      miMarceloReparteC= new MarceloReparteC(int.Parse(txtMarceloCodigo.Text));
			miMarceloReparteC.Show();	
		}

		private void grdTalla_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.Alt) 
				return;
			if (grdTalla.ActiveCell == null) return;
			switch(e.KeyValue)
			{
				case (int) Keys.PageDown:
					if(grdTalla.ActiveCell.DroppedDown == false)
						grdTalla.ActiveCell.DroppedDown = true;
					break;

				case (int) Keys.Up:
					if (grdTalla.ActiveCell.DroppedDown) return;

					grdTalla.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grdTalla.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					grdTalla.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (grdTalla.ActiveCell.DroppedDown) return;
			
					grdTalla.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grdTalla.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					grdTalla.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Enter:
					grdTalla.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grdTalla.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					grdTalla.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
			}		
		}
	}
}
