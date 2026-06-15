using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de MCodigos.
	/// </summary>
	public class MCodigos : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsMMarca;
		private C1.Data.C1DataView cdvMMarca;
		private C1.Data.C1DataSet cdsMColor;
		private C1.Data.C1DataView cdvMColor;
		private C1.Data.C1DataView cdvMSubtipo;
		private C1.Data.C1DataSet cdsMSubtipo;
		private C1.Data.C1DataSet cdsMTipo;
		private C1.Data.C1DataView cdvMTipo;
		private C1.Data.C1DataSet cdsMUsuario;
		private C1.Data.C1DataView cdvMUsuario;
		private DevExpress.XtraEditors.SimpleButton btSubGrupo;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private DevExpress.XtraEditors.SimpleButton simpleButton1;
		private System.Windows.Forms.Label label3;
		private DevExpress.XtraEditors.SimpleButton simpleButton2;
		private System.Windows.Forms.Label label4;
		private DevExpress.XtraEditors.SimpleButton simpleButton3;
		private System.Windows.Forms.Label label5;
		private DevExpress.XtraEditors.SimpleButton simpleButton4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarca;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubtipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbColor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbUsuario;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.Misc.UltraButton btIngresar;
		private Infragistics.Win.Misc.UltraButton btReparte;
		private Infragistics.Win.Misc.UltraButton btAlmacen;
		private C1.Data.C1DataSet cdsMarceloCodigo;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdTalla;
		private Infragistics.Win.Misc.UltraButton btRepartir;
		/// <summary>
		/// 
		C1.Data.C1DataRow drMCodigo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCosto;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantidad;
		private C1.Data.C1DataSet cdsTalla;
		private C1.Data.C1DataView cdvTallaH;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDesde;
		private C1.Data.C1DataView cdvTallaD;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmTallaH;
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
    float FCosto=0;
		private System.Windows.Forms.Label lblidMarceloCodigo;
		private Infragistics.Win.Misc.UltraButton btImprime;
		private System.Windows.Forms.CheckBox checkBox1;
		private C1.Data.C1DataView cdvReferencia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReferencia;
		private C1.Data.C1DataSet cdsMarceloCodigoref;
		private System.Windows.Forms.CheckBox checkBox2;
		private Infragistics.Win.Misc.UltraButton btImpresiones;
		float FPrecio=0;
		public MCodigos()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MCodigos));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloSubtipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloSubtipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtipo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloColor", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloMarca", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloUsuario", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloCodigo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloMaestro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla16");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla17");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla19");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla20");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla21");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla22");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla23");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla24");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla25");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla16");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla17");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla19");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla20");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla21");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla22");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla23");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla24");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla25");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unir");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloTalla", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTalla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Orden");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloTalla", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTalla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Orden");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloCodigo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloMaestro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla16");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla17");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla19");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla20");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla21");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla22");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla23");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla24");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla25");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla16");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla17");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla19");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla20");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla21");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla22");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla23");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla24");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calla25");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unir");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			this.cdsMMarca = new C1.Data.C1DataSet();
			this.cdvMMarca = new C1.Data.C1DataView();
			this.cdsMColor = new C1.Data.C1DataSet();
			this.cdvMColor = new C1.Data.C1DataView();
			this.cdvMSubtipo = new C1.Data.C1DataView();
			this.cdsMSubtipo = new C1.Data.C1DataSet();
			this.cdsMTipo = new C1.Data.C1DataSet();
			this.cdvMTipo = new C1.Data.C1DataView();
			this.cdsMUsuario = new C1.Data.C1DataSet();
			this.cdvMUsuario = new C1.Data.C1DataView();
			this.btSubGrupo = new DevExpress.XtraEditors.SimpleButton();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbSubtipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.simpleButton2 = new DevExpress.XtraEditors.SimpleButton();
			this.label4 = new System.Windows.Forms.Label();
			this.simpleButton3 = new DevExpress.XtraEditors.SimpleButton();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbColor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.simpleButton4 = new DevExpress.XtraEditors.SimpleButton();
			this.cmbMarca = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbUsuario = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsMarceloCodigo = new C1.Data.C1DataSet();
			this.label6 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btIngresar = new Infragistics.Win.Misc.UltraButton();
			this.btReparte = new Infragistics.Win.Misc.UltraButton();
			this.btAlmacen = new Infragistics.Win.Misc.UltraButton();
			this.grdTalla = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btRepartir = new Infragistics.Win.Misc.UltraButton();
			this.txtCosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtPrecio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtCantidad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.cdsTalla = new C1.Data.C1DataSet();
			this.cdvTallaH = new C1.Data.C1DataView();
			this.cmbDesde = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvTallaD = new C1.Data.C1DataView();
			this.cmTallaH = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblidMarceloCodigo = new System.Windows.Forms.Label();
			this.btImprime = new Infragistics.Win.Misc.UltraButton();
			this.checkBox1 = new System.Windows.Forms.CheckBox();
			this.cdvReferencia = new C1.Data.C1DataView();
			this.cdsMarceloCodigoref = new C1.Data.C1DataSet();
			this.cmbReferencia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.checkBox2 = new System.Windows.Forms.CheckBox();
			this.btImpresiones = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsMMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMColor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMColor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMSubtipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMSubtipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubtipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdTalla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTalla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTallaH)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTallaD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmTallaH)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloCodigoref)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReferencia)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsMMarca
			// 
			this.cdsMMarca.BindingContextCtrl = this;
			this.cdsMMarca.DataLibrary = "LibFacturacion";
			this.cdsMMarca.DataLibraryUrl = "";
			this.cdsMMarca.DataSetDef = "dsMarceloMarca";
			this.cdsMMarca.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMMarca.Name = "cdsMMarca";
			this.cdsMMarca.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMMarca.SchemaDef = null;
			this.cdsMMarca.BeforeFill += new C1.Data.FillEventHandler(this.cdsMMarca_BeforeFill);
			// 
			// cdvMMarca
			// 
			this.cdvMMarca.BindingContextCtrl = this;
			this.cdvMMarca.DataSet = this.cdsMMarca;
			this.cdvMMarca.Sort = "marca";
			this.cdvMMarca.TableName = "";
			this.cdvMMarca.TableViewName = "MarceloMarca";
			// 
			// cdsMColor
			// 
			this.cdsMColor.BindingContextCtrl = this;
			this.cdsMColor.DataLibrary = "LibFacturacion";
			this.cdsMColor.DataLibraryUrl = "";
			this.cdsMColor.DataSetDef = "dsMarceloColor";
			this.cdsMColor.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMColor.Name = "cdsMColor";
			this.cdsMColor.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMColor.SchemaDef = null;
			this.cdsMColor.BeforeFill += new C1.Data.FillEventHandler(this.cdsMColor_BeforeFill);
			// 
			// cdvMColor
			// 
			this.cdvMColor.BindingContextCtrl = this;
			this.cdvMColor.DataSet = this.cdsMColor;
			this.cdvMColor.Sort = "color";
			this.cdvMColor.TableName = "";
			this.cdvMColor.TableViewName = "MarceloColor";
			// 
			// cdvMSubtipo
			// 
			this.cdvMSubtipo.BindingContextCtrl = this;
			this.cdvMSubtipo.DataSet = this.cdsMSubtipo;
			this.cdvMSubtipo.Sort = "subtipo";
			this.cdvMSubtipo.TableName = "";
			this.cdvMSubtipo.TableViewName = "MarceloSubtipo";
			// 
			// cdsMSubtipo
			// 
			this.cdsMSubtipo.BindingContextCtrl = this;
			this.cdsMSubtipo.DataLibrary = "LibFacturacion";
			this.cdsMSubtipo.DataLibraryUrl = "";
			this.cdsMSubtipo.DataSetDef = "dsMarceloSubtipo";
			this.cdsMSubtipo.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMSubtipo.Name = "cdsMSubtipo";
			this.cdsMSubtipo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMSubtipo.SchemaDef = null;
			this.cdsMSubtipo.BeforeFill += new C1.Data.FillEventHandler(this.cdsMSubtipo_BeforeFill);
			// 
			// cdsMTipo
			// 
			this.cdsMTipo.BindingContextCtrl = this;
			this.cdsMTipo.DataLibrary = "LibFacturacion";
			this.cdsMTipo.DataLibraryUrl = "";
			this.cdsMTipo.DataSetDef = "dsMarceloTipo";
			this.cdsMTipo.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMTipo.Name = "cdsMTipo";
			this.cdsMTipo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMTipo.SchemaDef = null;
			this.cdsMTipo.BeforeFill += new C1.Data.FillEventHandler(this.cdsMTipo_BeforeFill);
			// 
			// cdvMTipo
			// 
			this.cdvMTipo.BindingContextCtrl = this;
			this.cdvMTipo.DataSet = this.cdsMTipo;
			this.cdvMTipo.Sort = "tipo";
			this.cdvMTipo.TableName = "";
			this.cdvMTipo.TableViewName = "MarceloTipo";
			// 
			// cdsMUsuario
			// 
			this.cdsMUsuario.BindingContextCtrl = this;
			this.cdsMUsuario.DataLibrary = "LibFacturacion";
			this.cdsMUsuario.DataLibraryUrl = "";
			this.cdsMUsuario.DataSetDef = "dsMarceloUsuario";
			this.cdsMUsuario.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMUsuario.Name = "cdsMUsuario";
			this.cdsMUsuario.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMUsuario.SchemaDef = null;
			this.cdsMUsuario.BeforeFill += new C1.Data.FillEventHandler(this.cdsMUsuario_BeforeFill);
			// 
			// cdvMUsuario
			// 
			this.cdvMUsuario.BindingContextCtrl = this;
			this.cdvMUsuario.DataSet = this.cdsMUsuario;
			this.cdvMUsuario.TableName = "";
			this.cdvMUsuario.TableViewName = "MarceloUsuario";
			// 
			// btSubGrupo
			// 
			this.btSubGrupo.Image = ((System.Drawing.Image)(resources.GetObject("btSubGrupo.Image")));
			this.btSubGrupo.Location = new System.Drawing.Point(296, 72);
			this.btSubGrupo.Name = "btSubGrupo";
			this.btSubGrupo.Size = new System.Drawing.Size(24, 23);
			this.btSubGrupo.TabIndex = 22;
			this.btSubGrupo.ToolTip = "Sebgrupo de artículos";
			this.btSubGrupo.Click += new System.EventHandler(this.btSubGrupo_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(48, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 25;
			this.label1.Text = "Marca:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(48, 104);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(29, 16);
			this.label2.TabIndex = 28;
			this.label2.Text = "Tipo:";
			// 
			// simpleButton1
			// 
			this.simpleButton1.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton1.Image")));
			this.simpleButton1.Location = new System.Drawing.Point(296, 104);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(24, 23);
			this.simpleButton1.TabIndex = 27;
			this.simpleButton1.ToolTip = "Sebgrupo de artículos";
			this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(48, 136);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(46, 16);
			this.label3.TabIndex = 31;
			this.label3.Text = "Subtipo:";
			// 
			// cmbSubtipo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubtipo.Appearance = appearance1;
			this.cmbSubtipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubtipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubtipo.DataMember = "MarceloSubtipo";
			this.cmbSubtipo.DataSource = this.cdsMSubtipo;
			this.cmbSubtipo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 61;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 107;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 204;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 174;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbSubtipo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSubtipo.DisplayMember = "Subtipo";
			this.cmbSubtipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubtipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubtipo.Location = new System.Drawing.Point(96, 136);
			this.cmbSubtipo.Name = "cmbSubtipo";
			this.cmbSubtipo.Size = new System.Drawing.Size(192, 21);
			this.cmbSubtipo.TabIndex = 2;
			this.cmbSubtipo.ValueMember = "Codigo";
			this.cmbSubtipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSubtipo_KeyDown);
			this.cmbSubtipo.Validated += new System.EventHandler(this.cmbSubtipo_Validated);
			// 
			// simpleButton2
			// 
			this.simpleButton2.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton2.Image")));
			this.simpleButton2.Location = new System.Drawing.Point(296, 136);
			this.simpleButton2.Name = "simpleButton2";
			this.simpleButton2.Size = new System.Drawing.Size(24, 23);
			this.simpleButton2.TabIndex = 30;
			this.simpleButton2.ToolTip = "Sebgrupo de artículos";
			this.simpleButton2.Click += new System.EventHandler(this.simpleButton2_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(368, 72);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(46, 16);
			this.label4.TabIndex = 34;
			this.label4.Text = "Usuario:";
			// 
			// simpleButton3
			// 
			this.simpleButton3.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton3.Image")));
			this.simpleButton3.Location = new System.Drawing.Point(632, 72);
			this.simpleButton3.Name = "simpleButton3";
			this.simpleButton3.Size = new System.Drawing.Size(24, 23);
			this.simpleButton3.TabIndex = 33;
			this.simpleButton3.ToolTip = "Sebgrupo de artículos";
			this.simpleButton3.Click += new System.EventHandler(this.simpleButton3_Click);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(368, 104);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(34, 16);
			this.label5.TabIndex = 37;
			this.label5.Text = "Color:";
			// 
			// cmbColor
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbColor.Appearance = appearance2;
			this.cmbColor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbColor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbColor.DataSource = this.cdvMColor;
			this.cmbColor.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 83;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 174;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbColor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbColor.DisplayMember = "Color";
			this.cmbColor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbColor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbColor.Location = new System.Drawing.Point(432, 104);
			this.cmbColor.Name = "cmbColor";
			this.cmbColor.Size = new System.Drawing.Size(192, 21);
			this.cmbColor.TabIndex = 4;
			this.cmbColor.ValueMember = "Codigo";
			this.cmbColor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbColor_KeyDown);
			this.cmbColor.Validated += new System.EventHandler(this.cmbColor_Validated);
			// 
			// simpleButton4
			// 
			this.simpleButton4.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton4.Image")));
			this.simpleButton4.Location = new System.Drawing.Point(632, 104);
			this.simpleButton4.Name = "simpleButton4";
			this.simpleButton4.Size = new System.Drawing.Size(24, 23);
			this.simpleButton4.TabIndex = 36;
			this.simpleButton4.ToolTip = "Sebgrupo de artículos";
			this.simpleButton4.Click += new System.EventHandler(this.simpleButton4_Click);
			// 
			// cmbMarca
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMarca.Appearance = appearance3;
			this.cmbMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMarca.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarca.DataSource = this.cdvMMarca;
			this.cmbMarca.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 85;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 63;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 111;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbMarca.DisplayMember = "Marca";
			this.cmbMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMarca.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMarca.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMarca.Location = new System.Drawing.Point(96, 72);
			this.cmbMarca.Name = "cmbMarca";
			this.cmbMarca.Size = new System.Drawing.Size(192, 21);
			this.cmbMarca.TabIndex = 0;
			this.cmbMarca.ValueMember = "Codigo";
			this.cmbMarca.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbMarca_KeyDown);
			this.cmbMarca.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbMarca_KeyPress);
			this.cmbMarca.Validated += new System.EventHandler(this.cmbMarca_Validated);
			this.cmbMarca.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbMarca_InitializeLayout);
			// 
			// cmbTipo
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance4;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.cdvMTipo;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 60;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Width = 174;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipo.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.AppearancesOnly;
			this.cmbTipo.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.AllRowsInBand;
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(96, 104);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(192, 21);
			this.cmbTipo.TabIndex = 1;
			this.cmbTipo.ValueMember = "idMarceloTipo";
			this.cmbTipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipo_KeyDown);
			this.cmbTipo.Validated += new System.EventHandler(this.cmbTipo_Validated);
			// 
			// cmbUsuario
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbUsuario.Appearance = appearance5;
			this.cmbUsuario.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbUsuario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUsuario.DataSource = this.cdvMUsuario;
			this.cmbUsuario.DisplayLayout.AutoFitColumns = true;
			ultraGridBand5.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 89;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Width = 174;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbUsuario.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbUsuario.DisplayMember = "Usuario";
			this.cmbUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbUsuario.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUsuario.Location = new System.Drawing.Point(432, 72);
			this.cmbUsuario.Name = "cmbUsuario";
			this.cmbUsuario.Size = new System.Drawing.Size(192, 21);
			this.cmbUsuario.TabIndex = 3;
			this.cmbUsuario.ValueMember = "Codigo";
			this.cmbUsuario.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbUsuario_KeyDown);
			this.cmbUsuario.Validated += new System.EventHandler(this.cmbUsuario_Validated);
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataSource = this.cdvMTipo;
			ultraGridBand6.AddButtonCaption = "Proyecto";
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Width = 213;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ultraDropDown1.DisplayMember = "Tipo";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(288, 8);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(232, 64);
			this.ultraDropDown1.TabIndex = 121;
			this.ultraDropDown1.Text = "a";
			this.ultraDropDown1.ValueMember = "idMarceloTipo";
			this.ultraDropDown1.Visible = false;
			// 
			// txtArticulo
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance6;
			this.txtArticulo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarceloCodigo, "MarceloCodigo.Descripcion"));
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Enabled = false;
			this.txtArticulo.Location = new System.Drawing.Point(152, 184);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(456, 21);
			this.txtArticulo.TabIndex = 122;
			// 
			// cdsMarceloCodigo
			// 
			this.cdsMarceloCodigo.BindingContextCtrl = this;
			this.cdsMarceloCodigo.DataLibrary = "LibFacturacion";
			this.cdsMarceloCodigo.DataLibraryUrl = "";
			this.cdsMarceloCodigo.DataSetDef = "MarceloCodigo";
			this.cdsMarceloCodigo.FillOnRequest = false;
			this.cdsMarceloCodigo.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMarceloCodigo.Name = "cdsMarceloCodigo";
			this.cdsMarceloCodigo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMarceloCodigo.SchemaDef = null;
			this.cdsMarceloCodigo.BeforeFill += new C1.Data.FillEventHandler(this.cdsMarceloCodigo_BeforeFill);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(368, 144);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(62, 16);
			this.label6.TabIndex = 124;
			this.label6.Text = "Referencia:";
			// 
			// txtCodigo
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance7;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarceloCodigo, "MarceloCodigo.Codigo"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(248, 160);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(240, 21);
			this.txtCodigo.TabIndex = 127;
			// 
			// btIngresar
			// 
			this.btIngresar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btIngresar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btIngresar.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btIngresar.Location = new System.Drawing.Point(352, 432);
			this.btIngresar.Name = "btIngresar";
			this.btIngresar.Size = new System.Drawing.Size(104, 24);
			this.btIngresar.TabIndex = 13;
			this.btIngresar.Text = "Ingresar";
			this.btIngresar.Click += new System.EventHandler(this.btIngresar_Click);
			// 
			// btReparte
			// 
			this.btReparte.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btReparte.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btReparte.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btReparte.Location = new System.Drawing.Point(184, 432);
			this.btReparte.Name = "btReparte";
			this.btReparte.Size = new System.Drawing.Size(56, 23);
			this.btReparte.TabIndex = 12;
			this.btReparte.Text = "Unir";
			this.btReparte.Click += new System.EventHandler(this.btReparte_Click);
			// 
			// btAlmacen
			// 
			this.btAlmacen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAlmacen.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAlmacen.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btAlmacen.Location = new System.Drawing.Point(24, 432);
			this.btAlmacen.Name = "btAlmacen";
			this.btAlmacen.Size = new System.Drawing.Size(80, 23);
			this.btAlmacen.TabIndex = 11;
			this.btAlmacen.Text = "Almacen";
			this.btAlmacen.Click += new System.EventHandler(this.btAlmacen_Click);
			// 
			// grdTalla
			// 
			this.grdTalla.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdTalla.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdTalla.DataMember = "MarceloCodigo";
			this.grdTalla.DataSource = this.cdsMarceloCodigo;
			appearance8.BackColor = System.Drawing.Color.White;
			this.grdTalla.DisplayLayout.Appearance = appearance8;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 13;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 13;
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridColumn22.Width = 110;
			ultraGridColumn23.Header.VisiblePosition = 3;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 13;
			ultraGridColumn24.Header.VisiblePosition = 4;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 10;
			ultraGridColumn25.Header.VisiblePosition = 5;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 10;
			ultraGridColumn26.Header.VisiblePosition = 6;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 8;
			ultraGridColumn27.Header.VisiblePosition = 7;
			ultraGridColumn27.Width = 36;
			ultraGridColumn28.Header.VisiblePosition = 8;
			ultraGridColumn28.Width = 36;
			ultraGridColumn29.Header.VisiblePosition = 9;
			ultraGridColumn29.Width = 36;
			ultraGridColumn30.Header.VisiblePosition = 10;
			ultraGridColumn30.Width = 36;
			ultraGridColumn31.Header.VisiblePosition = 11;
			ultraGridColumn31.Width = 36;
			ultraGridColumn32.Header.VisiblePosition = 12;
			ultraGridColumn32.Width = 36;
			ultraGridColumn33.Header.VisiblePosition = 13;
			ultraGridColumn33.Width = 36;
			ultraGridColumn34.Header.VisiblePosition = 14;
			ultraGridColumn34.Width = 36;
			ultraGridColumn35.Header.VisiblePosition = 15;
			ultraGridColumn35.Width = 36;
			ultraGridColumn36.Header.VisiblePosition = 16;
			ultraGridColumn36.Width = 36;
			ultraGridColumn37.Header.VisiblePosition = 17;
			ultraGridColumn37.Width = 36;
			ultraGridColumn38.Header.VisiblePosition = 18;
			ultraGridColumn38.Width = 36;
			ultraGridColumn39.Header.VisiblePosition = 19;
			ultraGridColumn39.Width = 36;
			ultraGridColumn40.Header.VisiblePosition = 20;
			ultraGridColumn40.Width = 36;
			ultraGridColumn41.Header.VisiblePosition = 21;
			ultraGridColumn41.Width = 36;
			ultraGridColumn42.Header.VisiblePosition = 22;
			ultraGridColumn42.Width = 36;
			ultraGridColumn43.Header.VisiblePosition = 23;
			ultraGridColumn43.Width = 36;
			ultraGridColumn44.Header.VisiblePosition = 24;
			ultraGridColumn44.Width = 36;
			ultraGridColumn45.Header.VisiblePosition = 25;
			ultraGridColumn45.Width = 36;
			ultraGridColumn46.Header.VisiblePosition = 26;
			ultraGridColumn46.Width = 36;
			ultraGridColumn47.Header.VisiblePosition = 27;
			ultraGridColumn47.Width = 36;
			ultraGridColumn48.Header.VisiblePosition = 28;
			ultraGridColumn48.Width = 36;
			ultraGridColumn49.Header.VisiblePosition = 29;
			ultraGridColumn49.Width = 36;
			ultraGridColumn50.Header.VisiblePosition = 30;
			ultraGridColumn50.Width = 36;
			ultraGridColumn51.Header.VisiblePosition = 31;
			ultraGridColumn51.Width = 36;
			ultraGridColumn52.Header.VisiblePosition = 32;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 13;
			ultraGridColumn53.Header.VisiblePosition = 33;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 13;
			ultraGridColumn54.Header.VisiblePosition = 34;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 13;
			ultraGridColumn55.Header.VisiblePosition = 35;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 13;
			ultraGridColumn56.Header.VisiblePosition = 36;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 13;
			ultraGridColumn57.Header.VisiblePosition = 37;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 13;
			ultraGridColumn58.Header.VisiblePosition = 38;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 13;
			ultraGridColumn59.Header.VisiblePosition = 39;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 13;
			ultraGridColumn60.Header.VisiblePosition = 40;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 13;
			ultraGridColumn61.Header.VisiblePosition = 41;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 13;
			ultraGridColumn62.Header.VisiblePosition = 42;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 13;
			ultraGridColumn63.Header.VisiblePosition = 43;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 13;
			ultraGridColumn64.Header.VisiblePosition = 44;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 13;
			ultraGridColumn65.Header.VisiblePosition = 45;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 13;
			ultraGridColumn66.Header.VisiblePosition = 46;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn66.Width = 13;
			ultraGridColumn67.Header.VisiblePosition = 47;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 13;
			ultraGridColumn68.Header.VisiblePosition = 48;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 13;
			ultraGridColumn69.Header.VisiblePosition = 49;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 13;
			ultraGridColumn70.Header.VisiblePosition = 50;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 13;
			ultraGridColumn71.Header.VisiblePosition = 51;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 13;
			ultraGridColumn72.Header.VisiblePosition = 52;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn72.Width = 13;
			ultraGridColumn73.Header.VisiblePosition = 53;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn73.Width = 13;
			ultraGridColumn74.Header.VisiblePosition = 54;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn74.Width = 13;
			ultraGridColumn75.Header.VisiblePosition = 55;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 13;
			ultraGridColumn76.Header.VisiblePosition = 56;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 13;
			ultraGridColumn77.Header.VisiblePosition = 57;
			ultraGridColumn78.Header.VisiblePosition = 58;
			ultraGridColumn79.Header.VisiblePosition = 59;
			ultraGridBand7.Columns.AddRange(new object[] {
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
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68,
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73,
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79});
			this.grdTalla.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.grdTalla.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdTalla.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdTalla.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.grdTalla.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdTalla.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.grdTalla.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdTalla.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdTalla.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.grdTalla.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdTalla.Location = new System.Drawing.Point(8, 304);
			this.grdTalla.Name = "grdTalla";
			this.grdTalla.Size = new System.Drawing.Size(664, 104);
			this.grdTalla.TabIndex = 10;
			this.grdTalla.Error += new Infragistics.Win.UltraWinGrid.ErrorEventHandler(this.grdTalla_Error);
			this.grdTalla.AfterRowUpdate += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.grdTalla_AfterRowUpdate);
			this.grdTalla.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdTalla_CellChange);
			this.grdTalla.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.grdTalla_KeyPress);
			this.grdTalla.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.grdTalla_BeforeCellUpdate);
			this.grdTalla.BeforeCellActivate += new Infragistics.Win.UltraWinGrid.CancelableCellEventHandler(this.grdTalla_BeforeCellActivate);
			this.grdTalla.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdTalla_KeyDown);
			this.grdTalla.KeyUp += new System.Windows.Forms.KeyEventHandler(this.grdTalla_KeyUp);
			this.grdTalla.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdTalla_AfterCellUpdate);
			this.grdTalla.BeforeCellDeactivate += new System.ComponentModel.CancelEventHandler(this.grdTalla_BeforeCellDeactivate);
			// 
			// btRepartir
			// 
			this.btRepartir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btRepartir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btRepartir.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btRepartir.Location = new System.Drawing.Point(592, 432);
			this.btRepartir.Name = "btRepartir";
			this.btRepartir.Size = new System.Drawing.Size(80, 23);
			this.btRepartir.TabIndex = 14;
			this.btRepartir.Text = "Repartir";
			this.btRepartir.Click += new System.EventHandler(this.btRepartir_Click);
			// 
			// txtCosto
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCosto.Appearance = appearance13;
			this.txtCosto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarceloCodigo, "MarceloCodigo.Costo"));
			this.txtCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCosto.FormatString = "#,##0.00";
			this.txtCosto.Location = new System.Drawing.Point(128, 224);
			this.txtCosto.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtCosto.Name = "txtCosto";
			this.txtCosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCosto.PromptChar = ' ';
			this.txtCosto.Size = new System.Drawing.Size(104, 26);
			this.txtCosto.TabIndex = 7;
			this.txtCosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCosto_KeyDown);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(72, 232);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(34, 16);
			this.label7.TabIndex = 134;
			this.label7.Text = "Costo";
			// 
			// txtPrecio
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecio.Appearance = appearance14;
			this.txtPrecio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarceloCodigo, "MarceloCodigo.Precio"));
			this.txtPrecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecio.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPrecio.FormatString = "#,##0.00";
			this.txtPrecio.Location = new System.Drawing.Point(336, 224);
			this.txtPrecio.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtPrecio.Name = "txtPrecio";
			this.txtPrecio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecio.PromptChar = ' ';
			this.txtPrecio.Size = new System.Drawing.Size(104, 26);
			this.txtPrecio.TabIndex = 8;
			this.txtPrecio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPrecio_KeyDown);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(280, 232);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(43, 16);
			this.label8.TabIndex = 136;
			this.label8.Text = "Precio :";
			// 
			// txtCantidad
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCantidad.Appearance = appearance15;
			this.txtCantidad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarceloCodigo, "MarceloCodigo.Cantidad"));
			this.txtCantidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantidad.Enabled = false;
			this.txtCantidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCantidad.FormatString = "#,##0";
			this.txtCantidad.Location = new System.Drawing.Point(552, 224);
			this.txtCantidad.MaskInput = "nnnnnnn";
			this.txtCantidad.Name = "txtCantidad";
			this.txtCantidad.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCantidad.PromptChar = ' ';
			this.txtCantidad.Size = new System.Drawing.Size(56, 26);
			this.txtCantidad.TabIndex = 137;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(488, 232);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(56, 16);
			this.label9.TabIndex = 138;
			this.label9.Text = "Cantidad :";
			// 
			// cdsTalla
			// 
			this.cdsTalla.BindingContextCtrl = this;
			this.cdsTalla.DataLibrary = "LibFacturacion";
			this.cdsTalla.DataLibraryUrl = "";
			this.cdsTalla.DataSetDef = "dsMarceloTalla";
			this.cdsTalla.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsTalla.Name = "cdsTalla";
			this.cdsTalla.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTalla.SchemaDef = null;
			this.cdsTalla.BeforeFill += new C1.Data.FillEventHandler(this.cdsTalla_BeforeFill);
			// 
			// cdvTallaH
			// 
			this.cdvTallaH.BindingContextCtrl = this;
			this.cdvTallaH.DataSet = this.cdsTalla;
			this.cdvTallaH.TableName = "";
			this.cdvTallaH.TableViewName = "MarceloTalla";
			// 
			// cmbDesde
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDesde.Appearance = appearance16;
			this.cmbDesde.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDesde.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDesde.DataSource = this.cdvTallaD;
			this.cmbDesde.DisplayLayout.AutoFitColumns = true;
			ultraGridBand8.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn80.Header.VisiblePosition = 0;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn80.Width = 11;
			ultraGridColumn81.Header.VisiblePosition = 1;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.Width = 16;
			ultraGridColumn82.Header.VisiblePosition = 2;
			ultraGridColumn82.Width = 78;
			ultraGridColumn83.Header.VisiblePosition = 3;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn83.Width = 32;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn80,
																										 ultraGridColumn81,
																										 ultraGridColumn82,
																										 ultraGridColumn83});
			this.cmbDesde.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbDesde.DisplayMember = "Talla";
			this.cmbDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDesde.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDesde.Location = new System.Drawing.Point(128, 264);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.Size = new System.Drawing.Size(96, 21);
			this.cmbDesde.TabIndex = 9;
			this.cmbDesde.ValueMember = "idTalla";
			this.cmbDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDesde_KeyDown);
			this.cmbDesde.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbDesde_MouseDown);
			this.cmbDesde.Validated += new System.EventHandler(this.cmbDesde_Validated);
			// 
			// cdvTallaD
			// 
			this.cdvTallaD.AllowAddNew = false;
			this.cdvTallaD.BindingContextCtrl = this;
			this.cdvTallaD.DataSet = this.cdsTalla;
			this.cdvTallaD.TableName = "";
			this.cdvTallaD.TableViewName = "MarceloTalla";
			// 
			// cmTallaH
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmTallaH.Appearance = appearance17;
			this.cmTallaH.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmTallaH.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmTallaH.DataSource = this.cdvTallaH;
			this.cmTallaH.DisplayLayout.AutoFitColumns = true;
			ultraGridBand9.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn84.Header.VisiblePosition = 0;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn84.Width = 11;
			ultraGridColumn85.Header.VisiblePosition = 1;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn85.Width = 16;
			ultraGridColumn86.Header.VisiblePosition = 2;
			ultraGridColumn86.Width = 70;
			ultraGridColumn87.Header.VisiblePosition = 3;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn87.Width = 29;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn84,
																										 ultraGridColumn85,
																										 ultraGridColumn86,
																										 ultraGridColumn87});
			this.cmTallaH.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmTallaH.DisplayMember = "Talla";
			this.cmTallaH.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmTallaH.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmTallaH.Location = new System.Drawing.Point(296, 264);
			this.cmTallaH.Name = "cmTallaH";
			this.cmTallaH.Size = new System.Drawing.Size(88, 21);
			this.cmTallaH.TabIndex = 10;
			this.cmTallaH.ValueMember = "idTalla";
			this.cmTallaH.Validated += new System.EventHandler(this.cmTallaH_Validated);
			// 
			// lblidMarceloCodigo
			// 
			this.lblidMarceloCodigo.AutoSize = true;
			this.lblidMarceloCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsMarceloCodigo, "MarceloCodigo.idMarceloCodigo"));
			this.lblidMarceloCodigo.Location = new System.Drawing.Point(24, 16);
			this.lblidMarceloCodigo.Name = "lblidMarceloCodigo";
			this.lblidMarceloCodigo.Size = new System.Drawing.Size(39, 16);
			this.lblidMarceloCodigo.TabIndex = 139;
			this.lblidMarceloCodigo.Text = "Marca:";
			// 
			// btImprime
			// 
			this.btImprime.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btImprime.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btImprime.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btImprime.Location = new System.Drawing.Point(528, 464);
			this.btImprime.Name = "btImprime";
			this.btImprime.Size = new System.Drawing.Size(144, 32);
			this.btImprime.TabIndex = 140;
			this.btImprime.Text = "Imprimir Codigos";
			this.btImprime.Click += new System.EventHandler(this.btImprime_Click);
			// 
			// checkBox1
			// 
			this.checkBox1.Location = new System.Drawing.Point(424, 264);
			this.checkBox1.Name = "checkBox1";
			this.checkBox1.Size = new System.Drawing.Size(128, 24);
			this.checkBox1.TabIndex = 141;
			this.checkBox1.Text = "Solo deseo imprimir";
			// 
			// cdvReferencia
			// 
			this.cdvReferencia.AllowAddNew = false;
			this.cdvReferencia.BindingContextCtrl = this;
			this.cdvReferencia.DataSet = this.cdsMarceloCodigoref;
			this.cdvReferencia.Sort = "referencia";
			this.cdvReferencia.TableName = "";
			this.cdvReferencia.TableViewName = "MarceloCodigo";
			// 
			// cdsMarceloCodigoref
			// 
			this.cdsMarceloCodigoref.BindingContextCtrl = this;
			this.cdsMarceloCodigoref.DataLibrary = "LibFacturacion";
			this.cdsMarceloCodigoref.DataLibraryUrl = "";
			this.cdsMarceloCodigoref.DataSetDef = "MarceloCodigo";
			this.cdsMarceloCodigoref.FillOnRequest = false;
			this.cdsMarceloCodigoref.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMarceloCodigoref.Name = "cdsMarceloCodigoref";
			this.cdsMarceloCodigoref.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMarceloCodigoref.SchemaDef = null;
			// 
			// cmbReferencia
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbReferencia.Appearance = appearance18;
			this.cmbReferencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReferencia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReferencia.DataSource = this.cdvReferencia;
			this.cmbReferencia.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn88.Header.VisiblePosition = 0;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn88.Width = 8;
			ultraGridColumn89.Header.VisiblePosition = 1;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Width = 8;
			ultraGridColumn90.Header.VisiblePosition = 2;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Width = 8;
			ultraGridColumn91.Header.VisiblePosition = 3;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 8;
			ultraGridColumn92.Header.VisiblePosition = 4;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn92.Width = 8;
			ultraGridColumn93.Header.VisiblePosition = 5;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn93.Width = 8;
			ultraGridColumn94.Header.VisiblePosition = 6;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Width = 8;
			ultraGridColumn95.Header.VisiblePosition = 7;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn95.Width = 8;
			ultraGridColumn96.Header.VisiblePosition = 8;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn96.Width = 8;
			ultraGridColumn97.Header.VisiblePosition = 9;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 8;
			ultraGridColumn98.Header.VisiblePosition = 10;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn98.Width = 8;
			ultraGridColumn99.Header.VisiblePosition = 11;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn99.Width = 8;
			ultraGridColumn100.Header.VisiblePosition = 12;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn100.Width = 8;
			ultraGridColumn101.Header.VisiblePosition = 13;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn101.Width = 8;
			ultraGridColumn102.Header.VisiblePosition = 14;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn102.Width = 8;
			ultraGridColumn103.Header.VisiblePosition = 15;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn103.Width = 8;
			ultraGridColumn104.Header.VisiblePosition = 16;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn104.Width = 8;
			ultraGridColumn105.Header.VisiblePosition = 17;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn105.Width = 8;
			ultraGridColumn106.Header.VisiblePosition = 18;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn106.Width = 8;
			ultraGridColumn107.Header.VisiblePosition = 19;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn107.Width = 8;
			ultraGridColumn108.Header.VisiblePosition = 20;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn108.Width = 8;
			ultraGridColumn109.Header.VisiblePosition = 21;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn109.Width = 8;
			ultraGridColumn110.Header.VisiblePosition = 22;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn110.Width = 8;
			ultraGridColumn111.Header.VisiblePosition = 23;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn111.Width = 8;
			ultraGridColumn112.Header.VisiblePosition = 24;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn112.Width = 8;
			ultraGridColumn113.Header.VisiblePosition = 25;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn113.Width = 8;
			ultraGridColumn114.Header.VisiblePosition = 26;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn114.Width = 8;
			ultraGridColumn115.Header.VisiblePosition = 27;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn115.Width = 8;
			ultraGridColumn116.Header.VisiblePosition = 28;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn116.Width = 8;
			ultraGridColumn117.Header.VisiblePosition = 29;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn117.Width = 8;
			ultraGridColumn118.Header.VisiblePosition = 30;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn118.Width = 8;
			ultraGridColumn119.Header.VisiblePosition = 31;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn119.Width = 8;
			ultraGridColumn120.Header.VisiblePosition = 32;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn120.Width = 8;
			ultraGridColumn121.Header.VisiblePosition = 33;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn121.Width = 8;
			ultraGridColumn122.Header.VisiblePosition = 34;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn122.Width = 8;
			ultraGridColumn123.Header.VisiblePosition = 35;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn123.Width = 8;
			ultraGridColumn124.Header.VisiblePosition = 36;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn124.Width = 8;
			ultraGridColumn125.Header.VisiblePosition = 37;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn125.Width = 8;
			ultraGridColumn126.Header.VisiblePosition = 38;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn126.Width = 8;
			ultraGridColumn127.Header.VisiblePosition = 39;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn127.Width = 8;
			ultraGridColumn128.Header.VisiblePosition = 40;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn128.Width = 8;
			ultraGridColumn129.Header.VisiblePosition = 41;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn129.Width = 8;
			ultraGridColumn130.Header.VisiblePosition = 42;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn130.Width = 8;
			ultraGridColumn131.Header.VisiblePosition = 43;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn131.Width = 8;
			ultraGridColumn132.Header.VisiblePosition = 44;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn132.Width = 8;
			ultraGridColumn133.Header.VisiblePosition = 45;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn133.Width = 8;
			ultraGridColumn134.Header.VisiblePosition = 46;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn134.Width = 8;
			ultraGridColumn135.Header.VisiblePosition = 47;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn135.Width = 8;
			ultraGridColumn136.Header.VisiblePosition = 48;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn136.Width = 8;
			ultraGridColumn137.Header.VisiblePosition = 49;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn137.Width = 8;
			ultraGridColumn138.Header.VisiblePosition = 50;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn138.Width = 8;
			ultraGridColumn139.Header.VisiblePosition = 51;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn139.Width = 8;
			ultraGridColumn140.Header.VisiblePosition = 52;
			ultraGridColumn140.Hidden = true;
			ultraGridColumn140.Width = 8;
			ultraGridColumn141.Header.VisiblePosition = 53;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn141.Width = 8;
			ultraGridColumn142.Header.VisiblePosition = 54;
			ultraGridColumn142.Hidden = true;
			ultraGridColumn142.Width = 8;
			ultraGridColumn143.Header.VisiblePosition = 55;
			ultraGridColumn143.Hidden = true;
			ultraGridColumn143.Width = 8;
			ultraGridColumn144.Header.VisiblePosition = 56;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn144.Width = 8;
			ultraGridColumn145.Header.VisiblePosition = 57;
			ultraGridColumn145.Hidden = true;
			ultraGridColumn145.Width = 16;
			ultraGridColumn146.Header.VisiblePosition = 58;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn146.Width = 8;
			ultraGridColumn147.Header.VisiblePosition = 59;
			ultraGridColumn147.Width = 174;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn88,
																											ultraGridColumn89,
																											ultraGridColumn90,
																											ultraGridColumn91,
																											ultraGridColumn92,
																											ultraGridColumn93,
																											ultraGridColumn94,
																											ultraGridColumn95,
																											ultraGridColumn96,
																											ultraGridColumn97,
																											ultraGridColumn98,
																											ultraGridColumn99,
																											ultraGridColumn100,
																											ultraGridColumn101,
																											ultraGridColumn102,
																											ultraGridColumn103,
																											ultraGridColumn104,
																											ultraGridColumn105,
																											ultraGridColumn106,
																											ultraGridColumn107,
																											ultraGridColumn108,
																											ultraGridColumn109,
																											ultraGridColumn110,
																											ultraGridColumn111,
																											ultraGridColumn112,
																											ultraGridColumn113,
																											ultraGridColumn114,
																											ultraGridColumn115,
																											ultraGridColumn116,
																											ultraGridColumn117,
																											ultraGridColumn118,
																											ultraGridColumn119,
																											ultraGridColumn120,
																											ultraGridColumn121,
																											ultraGridColumn122,
																											ultraGridColumn123,
																											ultraGridColumn124,
																											ultraGridColumn125,
																											ultraGridColumn126,
																											ultraGridColumn127,
																											ultraGridColumn128,
																											ultraGridColumn129,
																											ultraGridColumn130,
																											ultraGridColumn131,
																											ultraGridColumn132,
																											ultraGridColumn133,
																											ultraGridColumn134,
																											ultraGridColumn135,
																											ultraGridColumn136,
																											ultraGridColumn137,
																											ultraGridColumn138,
																											ultraGridColumn139,
																											ultraGridColumn140,
																											ultraGridColumn141,
																											ultraGridColumn142,
																											ultraGridColumn143,
																											ultraGridColumn144,
																											ultraGridColumn145,
																											ultraGridColumn146,
																											ultraGridColumn147});
			this.cmbReferencia.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbReferencia.DisplayMember = "Referencia";
			this.cmbReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbReferencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReferencia.Location = new System.Drawing.Point(432, 136);
			this.cmbReferencia.Name = "cmbReferencia";
			this.cmbReferencia.Size = new System.Drawing.Size(192, 21);
			this.cmbReferencia.TabIndex = 5;
			this.cmbReferencia.ValueMember = "Referencia";
			this.cmbReferencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbReferencia_KeyDown_1);
			this.cmbReferencia.Validated += new System.EventHandler(this.cmbReferencia_Validated);
			// 
			// checkBox2
			// 
			this.checkBox2.Location = new System.Drawing.Point(568, 264);
			this.checkBox2.Name = "checkBox2";
			this.checkBox2.Size = new System.Drawing.Size(80, 24);
			this.checkBox2.TabIndex = 142;
			this.checkBox2.Text = "Ingreso";
			// 
			// btImpresiones
			// 
			this.btImpresiones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btImpresiones.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btImpresiones.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btImpresiones.Location = new System.Drawing.Point(24, 472);
			this.btImpresiones.Name = "btImpresiones";
			this.btImpresiones.Size = new System.Drawing.Size(104, 23);
			this.btImpresiones.TabIndex = 143;
			this.btImpresiones.Text = "Impresiones";
			this.btImpresiones.Click += new System.EventHandler(this.btImpresiones_Click);
			// 
			// MCodigos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(696, 501);
			this.Controls.Add(this.btImpresiones);
			this.Controls.Add(this.checkBox2);
			this.Controls.Add(this.cmbReferencia);
			this.Controls.Add(this.checkBox1);
			this.Controls.Add(this.btImprime);
			this.Controls.Add(this.lblidMarceloCodigo);
			this.Controls.Add(this.cmTallaH);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.txtCantidad);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtPrecio);
			this.Controls.Add(this.txtCosto);
			this.Controls.Add(this.btRepartir);
			this.Controls.Add(this.grdTalla);
			this.Controls.Add(this.btAlmacen);
			this.Controls.Add(this.btReparte);
			this.Controls.Add(this.btIngresar);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.ultraDropDown1);
			this.Controls.Add(this.cmbUsuario);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbMarca);
			this.Controls.Add(this.cmbColor);
			this.Controls.Add(this.simpleButton4);
			this.Controls.Add(this.simpleButton3);
			this.Controls.Add(this.cmbSubtipo);
			this.Controls.Add(this.simpleButton2);
			this.Controls.Add(this.simpleButton1);
			this.Controls.Add(this.btSubGrupo);
			this.Name = "MCodigos";
			this.Text = "v";
			this.Load += new System.EventHandler(this.MCodigos_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsMMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMColor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMColor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMSubtipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMSubtipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubtipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdTalla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTalla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTallaH)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTallaD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmTallaH)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloCodigoref)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReferencia)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void MCodigos_Load(object sender, System.EventArgs e)
		{
			cmbMarca.Value="A7";
			cmbTipo.Value=1;
 			cmbSubtipo.Value="D0"; 
			cmbUsuario.Value="1";
			cmbColor.Value="F9"; 
	  
      drMCodigo= this.cdsMarceloCodigo.TableViews["MarceloCodigo"].AddNew();   

		}

		private void btSubGrupo_Click(object sender, System.EventArgs e)
		{
			using (MMarca miMMarca = new MMarca())
			{
				miMMarca.ShowDialog();
			}
			this.cdsMMarca.Clear();
			this.cdsMMarca.Fill();
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			
			using (MTipo miMTipo = new MTipo())
			{
				miMTipo.ShowDialog();
			}
			this.cdsMTipo.Clear();
			this.cdsMTipo.Fill();

		}

		private void simpleButton2_Click(object sender, System.EventArgs e)
		{
		
			using (MSubtipo miMSubtipo = new MSubtipo())
			{
				miMSubtipo.ShowDialog();
			}
			this.cdsMSubtipo.Clear();
			this.cdsMSubtipo.Fill();

		}

		private void simpleButton3_Click(object sender, System.EventArgs e)
		{
			using (MUsuario miMUsuario = new MUsuario())
			{
				miMUsuario.ShowDialog();
			}
			this.cdsMUsuario.Clear();
			this.cdsMUsuario.Fill();
		}

		private void simpleButton4_Click(object sender, System.EventArgs e)
		{
			using (MColor miMColor = new MColor())
			{
				miMColor.ShowDialog();
			}
			this.cdsMColor.Clear();
			this.cdsMColor.Fill();
			//cdvMColor. 
	//		this.cdvMColor.Refresh();

			cmbColor.Select();  
			
		}

		private void cdsMMarca_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMMarca.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsMColor_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMColor.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsMSubtipo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMSubtipo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsMTipo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMTipo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsMUsuario_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMUsuario.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void simpleButton5_Click(object sender, System.EventArgs e)
		{
		/*	txtCodigo.Text =  cmbMarca.Value.ToString() +cmbTipo.Value.ToString()+ cmbSubtipo.Value.ToString()
			+ cmbUsuario.Value.ToString() + cmbColor.Value.ToString()+ cmbReferencia.Text;  
			txtArticulo.Text = cmbMarca.Text  + " " + cmbSubtipo.Text + " " + cmbUsuario.Text + " "+ cmbColor.Text;   
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
			{
				 dr.Cells["Codigo"].Value = txtCodigo.Text;
  		}
*/
		}

		private void grdTallas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btIngresar_Click(object sender, System.EventArgs e)
		{
		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
			{
				dr.Cells["Calla1"].Value =dr.Cells["Talla1"].Column.Header.Caption;    
			  dr.Cells["Calla2"].Value =dr.Cells["Talla2"].Column.Header.Caption;    
				dr.Cells["Calla3"].Value =dr.Cells["Talla3"].Column.Header.Caption;
				dr.Cells["Calla4"].Value =dr.Cells["Talla4"].Column.Header.Caption;    
				dr.Cells["Calla5"].Value =dr.Cells["Talla5"].Column.Header.Caption;    
				dr.Cells["Calla6"].Value =dr.Cells["Talla6"].Column.Header.Caption; 
				dr.Cells["Calla7"].Value =dr.Cells["Talla7"].Column.Header.Caption;    
				dr.Cells["Calla8"].Value =dr.Cells["Talla8"].Column.Header.Caption;    
				dr.Cells["Calla9"].Value =dr.Cells["Talla9"].Column.Header.Caption; 
				dr.Cells["Calla10"].Value =dr.Cells["Talla10"].Column.Header.Caption;    
				dr.Cells["Calla11"].Value =dr.Cells["Talla11"].Column.Header.Caption;    
				dr.Cells["Calla12"].Value =dr.Cells["Talla12"].Column.Header.Caption;    
				dr.Cells["Calla13"].Value =dr.Cells["Talla13"].Column.Header.Caption;
				dr.Cells["Calla14"].Value =dr.Cells["Talla14"].Column.Header.Caption;    
				dr.Cells["Calla15"].Value =dr.Cells["Talla15"].Column.Header.Caption;    
				dr.Cells["Calla16"].Value =dr.Cells["Talla16"].Column.Header.Caption; 
				dr.Cells["Calla17"].Value =dr.Cells["Talla17"].Column.Header.Caption;    
				dr.Cells["Calla18"].Value =dr.Cells["Talla18"].Column.Header.Caption;    
				dr.Cells["Calla19"].Value =dr.Cells["Talla19"].Column.Header.Caption; 
				dr.Cells["Calla20"].Value =dr.Cells["Talla20"].Column.Header.Caption;    
				dr.Cells["Calla21"].Value =dr.Cells["Talla21"].Column.Header.Caption;    
				dr.Cells["Calla22"].Value =dr.Cells["Talla22"].Column.Header.Caption;
				dr.Cells["Calla23"].Value =dr.Cells["Talla23"].Column.Header.Caption;    
				dr.Cells["Calla24"].Value =dr.Cells["Talla24"].Column.Header.Caption;    
				dr.Cells["Calla25"].Value =dr.Cells["Talla25"].Column.Header.Caption;
         
			}
			cdsMarceloCodigo.Update(); 
			//exec marcelocoimprime1 69,'D2D01F9140','ROJOnaranja','1040','PVP $ 45','REEBOK'
			//exec marcelocoimprime1 68,'D2D01F9140','ROJO','1040','PVP $ 45','NEW BALANCE'
			string sql;
			if (checkBox2.Checked==false)
			{
				sql="exec marcelocoimprime1 " + lblidMarceloCodigo.Text +",'"+
				txtCodigo.Text + "','"+ cmbColor.Text + "','"  + cmbReferencia.Text 
				+ "','"+ "PVP $: "	+ txtPrecio.Value.ToString() + "','"+cmbMarca.Text+"'";
				Funcion.EjecutaSQL(cdsMarceloCodigo,sql);
			}
			string txcodigo ="";
			txcodigo=txtCodigo.Text; 
			int codigo = int.Parse(lblidMarceloCodigo.Text);
			FCosto=float.Parse( drMCodigo["Costo"].ToString());
			FPrecio=float.Parse( drMCodigo["Precio"].ToString());
			cdsMarceloCodigo.Clear();
			drMCodigo = cdsMarceloCodigo.TableViews["MarceloCodigo"].AddNew(); 
			drMCodigo["Costo"]=FCosto;
			drMCodigo["Precio"]=FPrecio;

			if (checkBox1.Checked==true   )//SOLO IMPRIME SE BORRA EL CODIGO INGRESADO
			{
					sql ="delete from marcelocodigo where idmarcelocodigo="+ codigo.ToString();
					Funcion.EjecutaSQL(cdsMarceloCodigo,sql);  
				  MessageBox.Show("Codigo listo para imprimir");
			}	
			else
			{
						MessageBox.Show("Codigo ingresado");
			}		
			//exec MARCELOMANCOD  110,'A7D01F91040',10
			sql="exec MARCELOMANCOD "+ codigo.ToString() + ",'"+  txcodigo.ToString() +	"', "+ 	FPrecio.ToString();  	 
			Funcion.EjecutaSQL(cdsMarceloCodigo,sql);  
		}



		private void btReparte_Click(object sender, System.EventArgs e)
		{
			MReparte miMReparte= new MReparte();
			miMReparte.Show();
		}



		private void btAlmacen_Click(object sender, System.EventArgs e)
		{
			MAlmacen miMAlmacen= new MAlmacen();
			miMAlmacen.Show();
		}

		private void btRepartir_Click(object sender, System.EventArgs e)
		{
			MMaestro miMMaestro= new MMaestro();
			miMMaestro.Show();		
		}

		private void cmbTipo_Validated(object sender, System.EventArgs e)
		{
			C1.Data.FilterCondition fcPadre = new 
			C1.Data.FilterCondition(this.cdsMSubtipo , "MarceloSubTipo", "idmarcelotipo="+cmbTipo.Value);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			this.cdsMSubtipo.Fill(fcTotal, false);


		}

		private void cmbMarca_Validated(object sender, System.EventArgs e)
		{
	
			string stFiltro = " ";
			stFiltro += " marca='" + cmbMarca.Value.ToString()+"'";
			C1.Data.FilterCondition fcPadre = new 
			C1.Data.FilterCondition(this.cdsMarceloCodigoref , "MarceloCodigo", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			this.cdsMarceloCodigoref.Fill(fcTotal, false);
			CreaTexto();
		}

		private void cmbSubtipo_Validated(object sender, System.EventArgs e)
		{
			CreaTexto();
		}

		private void cmbUsuario_Validated(object sender, System.EventArgs e)
		{
			CreaTexto();
		}

		private void cmbColor_Validated(object sender, System.EventArgs e)
		{
			CreaTexto();
		}

		private void cmbReferencia_Validated(object sender, System.EventArgs e)
		{
			CreaTexto(); 
		}
		private void CreaTexto()
		{
			string color1= cmbColor.Value.ToString();	
			string sql="select color from marcelocolor where codigo='"+ color1+ "'";
;			string color2 = Funcion.sEscalarSQL(cdsMarceloCodigo,sql );  

			drMCodigo["Codigo"] =  cmbMarca.Value.ToString() + cmbSubtipo.Value.ToString()
				+ cmbUsuario.Value.ToString() + cmbColor.Value.ToString()+ cmbReferencia.Text;  		
			drMCodigo["Descripcion"] = cmbMarca.Text.Trim()  + " " + cmbSubtipo.Text.Trim() + " "
				+ cmbUsuario.Text.Trim() + " "+ color2.ToString().Trim();   
			
  	}

		private void Suma()
		{
			  int t=0;
				//drMCodigo["Cantidad"]=     		     		
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
}
				drMCodigo["Cantidad"]=t;     		     		
	

/*
				(int)drMCodigo["Talla1"]+(int)drMCodigo["Talla2"]+(int)drMCodigo["Talla3"]     		     		
	      +(int)drMCodigo["Talla4"]+(int)drMCodigo["Talla5"]+(int)drMCodigo["Talla6"]     		     		
				+(int)drMCodigo["Talla7"]+(int)drMCodigo["Talla8"]+(int)drMCodigo["Talla9"]     		     		
				+(int)drMCodigo["Talla10"]+(int)drMCodigo["Talla11"]+(int)drMCodigo["Talla12"]     		     		
				+(int)drMCodigo["Talla13"]+(int)drMCodigo["Talla14"]+(int)drMCodigo["Talla15"]     		     		
				+(int)drMCodigo["Talla16"]+(int)drMCodigo["Talla17"]+(int)drMCodigo["Talla18"]     		     		
				+(int)drMCodigo["Talla19"]+(int)drMCodigo["Talla20"]+(int)drMCodigo["Talla21"]     		     		
				+(int)drMCodigo["Talla22"]+(int)drMCodigo["Talla23"]+(int)drMCodigo["Talla24"]     		     		
				+(int)drMCodigo["Talla25"];
*/
		}

		private void cdsMarceloCodigo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMarceloCodigo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void grdTalla_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			Suma();
		}

		private void grdTalla_AfterRowUpdate(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			Suma();
		}

		private void grdTalla_BeforeCellActivate(object sender, Infragistics.Win.UltraWinGrid.CancelableCellEventArgs e)
		{
			Suma();
		}

		private void grdTalla_BeforeCellDeactivate(object sender, System.ComponentModel.CancelEventArgs e)
		{
			//Suma();
		}

		private void grdTalla_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			//Suma();
		}

		private void grdTalla_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		Suma();
		}

		private void grdTalla_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		//	Suma();
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
			Suma();
		}

		private void grdTalla_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			//Suma();
		}

		private void grdTalla_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		//Suma();
		}

		private void cdsTalla_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTalla.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cmbDesde_Validated(object sender, System.EventArgs e)
		{
				
		}

		private void rellena()
		{

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow d in grdTalla.Rows.All)
			{
				d.Cells["Talla1"].Value =0;   
				d.Cells["Talla2"].Value =0;    
				d.Cells["Talla3"].Value =0;
				d.Cells["Talla4"].Value =0;    
				d.Cells["Talla5"].Value =0;    
				d.Cells["Talla6"].Value =0; 
				d.Cells["Talla7"].Value =0;    
				d.Cells["Talla8"].Value =0;    
				d.Cells["Talla9"].Value =0; 
				d.Cells["Talla10"].Value =0;    
				d.Cells["Talla11"].Value =0;    
				d.Cells["Talla12"].Value =0;    
				d.Cells["Talla13"].Value =0;
				d.Cells["Talla14"].Value =0;    
				d.Cells["Talla15"].Value =0;    
				d.Cells["Talla16"].Value =0; 
				d.Cells["Talla17"].Value =0;    
				d.Cells["Talla18"].Value =0;    
				d.Cells["Talla19"].Value =0; 
				d.Cells["Talla20"].Value =0;    
				d.Cells["Talla21"].Value =0;    
				d.Cells["Talla22"].Value =0;
				d.Cells["Talla23"].Value =0;    
				d.Cells["Talla24"].Value =0;    
				d.Cells["Talla25"].Value =0;

				grdTalla.DisplayLayout.Bands[0].Columns["Talla1"].Header.Caption = "Talla1" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla2"].Header.Caption = "Talla2" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla3"].Header.Caption = "Talla3" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla4"].Header.Caption = "Talla4" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla5"].Header.Caption = "Talla5" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla6"].Header.Caption = "Talla6" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla7"].Header.Caption = "Talla7" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla8"].Header.Caption = "Talla8" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla9"].Header.Caption = "Talla9" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla10"].Header.Caption = "Talla10" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla11"].Header.Caption = "Talla11" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla12"].Header.Caption = "Talla12" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla13"].Header.Caption = "Talla13" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla14"].Header.Caption = "Talla14" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla15"].Header.Caption = "Talla15" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla16"].Header.Caption = "Talla16" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla17"].Header.Caption = "Talla17" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla18"].Header.Caption = "Talla18" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla19"].Header.Caption = "Talla19" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla20"].Header.Caption = "Talla20" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla21"].Header.Caption = "Talla21" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla22"].Header.Caption = "Talla22" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla23"].Header.Caption = "Talla23" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla24"].Header.Caption = "Talla24" ;
				grdTalla.DisplayLayout.Bands[0].Columns["Talla25"].Header.Caption = "Talla25" ;
				
			}


			foreach(Infragistics.Win.UltraWinGrid.UltraGridColumn dc in grdTalla.DisplayLayout.Bands[0].Columns)
			{
					
				dc.Hidden = true;
			}
			grdTalla.DisplayLayout.Bands[0].Columns["Codigo"].Hidden = false;
			grdTalla.DisplayLayout.Bands[0].Columns["Codigo"].Header.VisiblePosition = 1;
			grdTalla.DisplayLayout.Bands[0].Columns["Codigo"].Header.Caption = "Codigo" ;
			grdTalla.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;


			string stColumna = "";
			string stTalla="";
			int iPos=1;
			
			string stLista = "Select talla From marcelotalla  WHERE IDTALLA>="+ cmbDesde.Value
				+ " AND IDTALLa<=" + cmTallaH.Value   ;  

			SqlDataReader dr = Funcion.rEscalarSQL(cdsMColor, stLista, true);
			while(dr.Read())
				#region Lee
			{
				stTalla = dr.GetString(0);
				stColumna = "talla" + iPos.ToString().Trim();
				grdTalla.DisplayLayout.Bands[0].Columns[stColumna].Hidden = false;
				grdTalla.DisplayLayout.Bands[0].Columns[stColumna].Header.VisiblePosition = iPos+1;
				grdTalla.DisplayLayout.Bands[0].Columns[stColumna].Header.Caption = stTalla ;
				grdTalla.DisplayLayout.Bands[0].Columns[stColumna].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				iPos =iPos +1;
			}
			#endregion Lee
			dr.Close();
		}

		private void cmTallaH_Validated(object sender, System.EventArgs e)
		{
			rellena(); 

		}

		private void txtCosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
				this.txtPrecio.SelectAll(); 
}

		private void cmbReferencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		}

		private void txtPrecio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
				this.cmbDesde.Select(); 

		}

		private void cmbDesde_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			  
		}

		private void grdTalla_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e)
		{
			if (e.MaskErrorInfo == null)
			{
				e.Cancel = true; 
				MessageBox.Show(e.DataErrorInfo.ErrorText, "Informacion de Error");
				e.DataErrorInfo.Cell.Value = 0;

			}
		}

		private void btImprime_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process.Start("ZebraTest.exe");     
			   

		}

		private void cmbMarca_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbReferencia_Validated_1(object sender, System.EventArgs e)
		{
			CreaTexto();


		}

		private void cmbMarca_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{

		}

		private void cmbMarca_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
				this.cmbTipo.Select(); 
		}

		private void cmbTipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
				this.cmbSubtipo.Select(); 
		}

		private void cmbSubtipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
				this.cmbUsuario.Select(); 
		}

		private void cmbUsuario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
				this.cmbColor.Select();		
		}

		private void cmbColor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
				this.cmbReferencia.Select(); 
		}

		private void cmbReferencia_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
				this.txtCosto.SelectAll(); 
		}

		private void cmbDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
				this.cmTallaH.Select(); 
		}

		private void btImpresiones_Click(object sender, System.EventArgs e)
		{
			MMImprimir  miMMImprimir= new MMImprimir();
			miMMImprimir.Show();
		}
	}
}
