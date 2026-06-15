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
	/// Descripción breve de ArtGrupo.
	/// </summary>
	public class ArtGrupo : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsGrupoArt;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCuenta;
		private C1.Data.C1DataView cdvCuentaNombre;
		private C1.Data.C1DataView cdvTipo;
		private System.Windows.Forms.Label label6;
		private C1.Data.C1DataSet cdsArticulo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdGrupoArt;
		private C1.Data.C1DataView cdvGrupoArt;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCompraC;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVentaC;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCostoC;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCostoD;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVentaD;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCompraD;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.Windows.Forms.Label lblCompra;
		private System.Windows.Forms.Label lblVenta;
		private System.Windows.Forms.Label lblCosto;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDescD;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDescC;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDevolucionD;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDevolucionC;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticuloSRI;
		private DevExpress.XtraEditors.SimpleButton btGenerar;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEgreso;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTransfCod;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTransf;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.Misc.UltraButton btAuditoria;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnIncr1;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnIncr2;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnIncr3;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnIncr4;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnIncr5;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnIncr6;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label20;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.Misc.UltraButton btnIncr;
		private System.Windows.Forms.Label label21;
		private System.ComponentModel.IContainer components;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.Misc.UltraButton btSubGrupo;
		private Infragistics.Win.Misc.UltraButton btCategorias;
		private Infragistics.Win.Misc.UltraButton btMarcas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscaSubGrupo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscarCategoria;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscarMarca;

		C1.Data.C1DataRow drGrupo;

    public ArtGrupo()
		{
			InitializeComponent();
		}

		int IdGrupo = 0;
		int IInventario = 0; //0 NO FILTRA -- 1 INVENTARIO -- 2 NO INVENTARIO
		public ArtGrupo(int idGrupo)
		{
			InitializeComponent();
			IdGrupo = idGrupo;
		}

		public ArtGrupo(int idGrupo, int iInventario)
		{
			InitializeComponent();
			IdGrupo = idGrupo;
			IInventario = iInventario;
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloSRI", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(ArtGrupo));
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IncrPI");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IncrPA");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IncrPI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IncrPA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMarca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand19 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Categoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCategoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Categoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			this.cdsGrupoArt = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.lblCompra = new System.Windows.Forms.Label();
			this.lblVenta = new System.Windows.Forms.Label();
			this.lblCosto = new System.Windows.Forms.Label();
			this.cdvCuentaNombre = new C1.Data.C1DataView();
			this.cdvTipo = new C1.Data.C1DataView();
			this.label6 = new System.Windows.Forms.Label();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.txtIdGrupoArt = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdvGrupoArt = new C1.Data.C1DataView();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCompraC = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbVentaC = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCostoC = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCostoD = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbVentaD = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCompraD = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cmbDescD = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDescC = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbDevolucionD = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDevolucionC = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbArticuloSRI = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.ultraCombo1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbEgreso = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.label7 = new System.Windows.Forms.Label();
			this.cmbTransfCod = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTransf = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btAuditoria = new Infragistics.Win.Misc.UltraButton();
			this.label8 = new System.Windows.Forms.Label();
			this.spnIncr1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.spnIncr2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.spnIncr3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.spnIncr4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.spnIncr5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.spnIncr6 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnIncr = new Infragistics.Win.Misc.UltraButton();
			this.label21 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btSubGrupo = new Infragistics.Win.Misc.UltraButton();
			this.btCategorias = new Infragistics.Win.Misc.UltraButton();
			this.btMarcas = new Infragistics.Win.Misc.UltraButton();
			this.txtBuscaSubGrupo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtBuscarCategoria = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtBuscarMarca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCompraC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVentaC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCostoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCostoD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVentaD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCompraD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDescD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDescC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDevolucionD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDevolucionC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEgreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransfCod)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransf)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscaSubGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscarCategoria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscarMarca)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsGrupoArt
			// 
			this.cdsGrupoArt.BindingContextCtrl = this;
			this.cdsGrupoArt.CaseSensitive = false;
			this.cdsGrupoArt.DataLibrary = "LibFacturacion";
			this.cdsGrupoArt.DataLibraryUrl = "";
			this.cdsGrupoArt.DataSetDef = "dsArticuloTabla";
			this.cdsGrupoArt.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupoArt.Name = "cdsGrupoArt";
			this.cdsGrupoArt.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupoArt.SchemaDef = null;
			this.cdsGrupoArt.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupoArt_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Nombre:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(312, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(42, 16);
			this.label2.TabIndex = 4;
			this.label2.Text = "Código:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.RowFilter = "Grupo =0";
			this.cdvCuenta.Sort = "Codigo";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.CaseSensitive = false;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// lblCompra
			// 
			this.lblCompra.AutoSize = true;
			this.lblCompra.Location = new System.Drawing.Point(16, 400);
			this.lblCompra.Name = "lblCompra";
			this.lblCompra.Size = new System.Drawing.Size(47, 16);
			this.lblCompra.TabIndex = 6;
			this.lblCompra.Text = "Compra:";
			this.lblCompra.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblVenta
			// 
			this.lblVenta.AutoSize = true;
			this.lblVenta.Location = new System.Drawing.Point(16, 424);
			this.lblVenta.Name = "lblVenta";
			this.lblVenta.Size = new System.Drawing.Size(37, 16);
			this.lblVenta.TabIndex = 8;
			this.lblVenta.Text = "Venta:";
			this.lblVenta.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCosto
			// 
			this.lblCosto.AutoSize = true;
			this.lblCosto.Location = new System.Drawing.Point(16, 448);
			this.lblCosto.Name = "lblCosto";
			this.lblCosto.Size = new System.Drawing.Size(36, 16);
			this.lblCosto.TabIndex = 10;
			this.lblCosto.Text = "Costo:";
			this.lblCosto.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cdvCuentaNombre
			// 
			this.cdvCuentaNombre.BindingContextCtrl = this;
			this.cdvCuentaNombre.DataSet = this.cdsCuenta;
			this.cdvCuentaNombre.RowFilter = "Grupo = 0";
			this.cdvCuentaNombre.Sort = "Descripcion";
			this.cdvCuentaNombre.TableName = "";
			this.cdvCuentaNombre.TableViewName = "Cuenta";
			// 
			// cdvTipo
			// 
			this.cdvTipo.BindingContextCtrl = this;
			this.cdvTipo.DataSet = this.cdsGrupoArt;
			this.cdvTipo.TableName = "";
			this.cdvTipo.TableViewName = "ArticuloTipo";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 32);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(30, 16);
			this.label6.TabIndex = 16;
			this.label6.Text = "Tipo:";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulo";
			this.cdsArticulo.FillOnRequest = false;
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-ES");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// txtIdGrupoArt
			// 
			this.txtIdGrupoArt.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idGrupoArticulo"));
			this.txtIdGrupoArt.Location = new System.Drawing.Point(672, 8);
			this.txtIdGrupoArt.Name = "txtIdGrupoArt";
			this.txtIdGrupoArt.Size = new System.Drawing.Size(100, 22);
			this.txtIdGrupoArt.TabIndex = 18;
			// 
			// cdvGrupoArt
			// 
			this.cdvGrupoArt.BindingContextCtrl = this;
			this.cdvGrupoArt.DataSet = this.cdsGrupoArt;
			this.cdvGrupoArt.Sort = "idGrupoArticulo";
			this.cdvGrupoArt.TableName = "";
			this.cdvGrupoArt.TableViewName = "ArticuloGrupo";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// txtNombre
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance1;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.Grupo"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(96, 8);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(200, 22);
			this.txtNombre.TabIndex = 1;
			// 
			// txtCodigo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance2;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.Codigo"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(392, 8);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(176, 22);
			this.txtCodigo.TabIndex = 2;
			// 
			// cmbCompraC
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCompraC.Appearance = appearance3;
			this.cmbCompraC.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCompraC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCompraC.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas2"));
			this.cmbCompraC.DataSource = this.cdvCuenta;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 4;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 232;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbCompraC.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCompraC.DisplayMember = "CodRapido";
			this.cmbCompraC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCompraC.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCompraC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCompraC.Location = new System.Drawing.Point(96, 400);
			this.cmbCompraC.Name = "cmbCompraC";
			this.cmbCompraC.Size = new System.Drawing.Size(200, 21);
			this.cmbCompraC.TabIndex = 6;
			this.cmbCompraC.ValueMember = "idCuenta";
			this.cmbCompraC.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCompraD_MouseDown);
			this.cmbCompraC.Enter += new System.EventHandler(this.cmbCompraC_Enter);
			// 
			// cmbVentaC
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVentaC.Appearance = appearance4;
			this.cmbVentaC.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVentaC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVentaC.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas3"));
			this.cmbVentaC.DataSource = this.cdvCuenta;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 225;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbVentaC.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVentaC.DisplayMember = "CodRapido";
			this.cmbVentaC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVentaC.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVentaC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVentaC.Location = new System.Drawing.Point(96, 424);
			this.cmbVentaC.Name = "cmbVentaC";
			this.cmbVentaC.Size = new System.Drawing.Size(200, 21);
			this.cmbVentaC.TabIndex = 8;
			this.cmbVentaC.ValueMember = "idCuenta";
			this.cmbVentaC.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbVentaC_InitializeLayout);
			this.cmbVentaC.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbVentaD_MouseDown);
			// 
			// cmbCostoC
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCostoC.Appearance = appearance5;
			this.cmbCostoC.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCostoC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCostoC.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas4"));
			this.cmbCostoC.DataSource = this.cdvCuenta;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 239;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbCostoC.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCostoC.DisplayMember = "CodRapido";
			this.cmbCostoC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCostoC.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCostoC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCostoC.Location = new System.Drawing.Point(96, 448);
			this.cmbCostoC.Name = "cmbCostoC";
			this.cmbCostoC.Size = new System.Drawing.Size(200, 21);
			this.cmbCostoC.TabIndex = 10;
			this.cmbCostoC.ValueMember = "idCuenta";
			this.cmbCostoC.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCostoD_MouseDown);
			// 
			// cmbCostoD
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCostoD.Appearance = appearance6;
			this.cmbCostoD.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCostoD.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCostoD.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas4"));
			this.cmbCostoD.DataSource = this.cdvCuentaNombre;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn17.Header.VisiblePosition = 4;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Width = 293;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbCostoD.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCostoD.DisplayMember = "Descripcion";
			this.cmbCostoD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCostoD.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCostoD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCostoD.Location = new System.Drawing.Point(312, 448);
			this.cmbCostoD.Name = "cmbCostoD";
			this.cmbCostoD.Size = new System.Drawing.Size(200, 21);
			this.cmbCostoD.TabIndex = 11;
			this.cmbCostoD.ValueMember = "idCuenta";
			this.cmbCostoD.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCostoD_MouseDown);
			// 
			// cmbVentaD
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVentaD.Appearance = appearance7;
			this.cmbVentaD.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVentaD.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVentaD.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas3"));
			this.cmbVentaD.DataSource = this.cdvCuentaNombre;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn22.Header.VisiblePosition = 4;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 3;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Width = 290;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbVentaD.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbVentaD.DisplayMember = "Descripcion";
			this.cmbVentaD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVentaD.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVentaD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVentaD.Location = new System.Drawing.Point(312, 424);
			this.cmbVentaD.Name = "cmbVentaD";
			this.cmbVentaD.Size = new System.Drawing.Size(200, 21);
			this.cmbVentaD.TabIndex = 9;
			this.cmbVentaD.ValueMember = "idCuenta";
			this.cmbVentaD.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbVentaD_MouseDown);
			// 
			// cmbCompraD
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCompraD.Appearance = appearance8;
			this.cmbCompraD.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCompraD.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCompraD.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas2"));
			this.cmbCompraD.DataSource = this.cdvCuentaNombre;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn27.Header.VisiblePosition = 4;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 3;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Width = 279;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.cmbCompraD.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCompraD.DisplayMember = "Descripcion";
			this.cmbCompraD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCompraD.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCompraD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCompraD.Location = new System.Drawing.Point(312, 400);
			this.cmbCompraD.Name = "cmbCompraD";
			this.cmbCompraD.Size = new System.Drawing.Size(200, 21);
			this.cmbCompraD.TabIndex = 7;
			this.cmbCompraD.ValueMember = "idCuenta";
			this.cmbCompraD.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbCompraD_MouseDown);
			// 
			// cmbTipo
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance9;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idTipoGrupo"));
			this.cmbTipo.DataSource = this.cdvTipo;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 254;
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Width = 181;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(96, 32);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(200, 21);
			this.cmbTipo.TabIndex = 3;
			this.cmbTipo.ValueMember = "idTipoGrupo";
			this.cmbTipo.RowSelected += new Infragistics.Win.UltraWinGrid.RowSelectedEventHandler(this.cmbTipo_RowSelected);
			this.cmbTipo.ValueChanged += new System.EventHandler(this.cmbTipo_ValueChanged);
			this.cmbTipo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipo_InitializeLayout);
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "ArticuloGrupo";
			this.barraDato1.DataNombreId = "idGrupoArticulo";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsGrupoArt;
			this.barraDato1.DataTabla = "ArticuloGrupo";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(8, 592);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(280, 26);
			this.barraDato1.TabIndex = 0;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cmbDescD
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDescD.Appearance = appearance10;
			this.cmbDescD.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDescD.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDescD.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas5"));
			this.cmbDescD.DataSource = this.cdvCuentaNombre;
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridColumn34.Header.VisiblePosition = 4;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 3;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 2;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 0;
			ultraGridColumn37.Width = 293;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			this.cmbDescD.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbDescD.DisplayMember = "Descripcion";
			this.cmbDescD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDescD.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDescD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDescD.Location = new System.Drawing.Point(312, 472);
			this.cmbDescD.Name = "cmbDescD";
			this.cmbDescD.Size = new System.Drawing.Size(200, 21);
			this.cmbDescD.TabIndex = 13;
			this.cmbDescD.ValueMember = "idCuenta";
			this.cmbDescD.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbDescD_MouseDown);
			// 
			// cmbDescC
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDescC.Appearance = appearance11;
			this.cmbDescC.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDescC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDescC.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas5"));
			this.cmbDescC.DataSource = this.cdvCuenta;
			ultraGridColumn38.Header.VisiblePosition = 2;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 4;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 3;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 0;
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.Width = 239;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42});
			this.cmbDescC.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbDescC.DisplayMember = "CodRapido";
			this.cmbDescC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDescC.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDescC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDescC.Location = new System.Drawing.Point(96, 472);
			this.cmbDescC.Name = "cmbDescC";
			this.cmbDescC.Size = new System.Drawing.Size(200, 21);
			this.cmbDescC.TabIndex = 12;
			this.cmbDescC.ValueMember = "idCuenta";
			this.cmbDescC.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbDescD_MouseDown);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 472);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(57, 16);
			this.label3.TabIndex = 22;
			this.label3.Text = "Descuento";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDevolucionD
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDevolucionD.Appearance = appearance12;
			this.cmbDevolucionD.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDevolucionD.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDevolucionD.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas6"));
			this.cmbDevolucionD.DataSource = this.cdvCuentaNombre;
			ultraGridColumn43.Header.VisiblePosition = 1;
			ultraGridColumn44.Header.VisiblePosition = 4;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 3;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 2;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn47.Width = 293;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn43,
																											ultraGridColumn44,
																											ultraGridColumn45,
																											ultraGridColumn46,
																											ultraGridColumn47});
			this.cmbDevolucionD.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbDevolucionD.DisplayMember = "Descripcion";
			this.cmbDevolucionD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDevolucionD.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDevolucionD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDevolucionD.Location = new System.Drawing.Point(312, 496);
			this.cmbDevolucionD.Name = "cmbDevolucionD";
			this.cmbDevolucionD.Size = new System.Drawing.Size(200, 21);
			this.cmbDevolucionD.TabIndex = 15;
			this.cmbDevolucionD.ValueMember = "idCuenta";
			this.cmbDevolucionD.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbDevolucionD_MouseDown);
			// 
			// cmbDevolucionC
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDevolucionC.Appearance = appearance13;
			this.cmbDevolucionC.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDevolucionC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDevolucionC.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas6"));
			this.cmbDevolucionC.DataSource = this.cdvCuenta;
			ultraGridColumn48.Header.VisiblePosition = 2;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 4;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 3;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 0;
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Width = 239;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn48,
																											ultraGridColumn49,
																											ultraGridColumn50,
																											ultraGridColumn51,
																											ultraGridColumn52});
			this.cmbDevolucionC.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbDevolucionC.DisplayMember = "CodRapido";
			this.cmbDevolucionC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDevolucionC.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDevolucionC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDevolucionC.Location = new System.Drawing.Point(96, 496);
			this.cmbDevolucionC.Name = "cmbDevolucionC";
			this.cmbDevolucionC.Size = new System.Drawing.Size(200, 21);
			this.cmbDevolucionC.TabIndex = 14;
			this.cmbDevolucionC.ValueMember = "idCuenta";
			this.cmbDevolucionC.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbDevolucionD_MouseDown);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 496);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(59, 16);
			this.label4.TabIndex = 25;
			this.label4.Text = "Devolucion";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbArticuloSRI
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArticuloSRI.Appearance = appearance14;
			this.cmbArticuloSRI.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticuloSRI.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticuloSRI.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idArticuloSRI"));
			this.cmbArticuloSRI.DataMember = "ArticuloSRI";
			this.cmbArticuloSRI.DataSource = this.cdsGrupoArt;
			ultraGridBand12.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn53.Header.VisiblePosition = 0;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 113;
			ultraGridColumn54.Header.VisiblePosition = 1;
			ultraGridColumn54.Width = 323;
			ultraGridColumn55.Header.VisiblePosition = 3;
			ultraGridColumn55.Width = 57;
			ultraGridColumn56.Header.VisiblePosition = 4;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 5;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 2;
			ultraGridColumn58.Width = 73;
			ultraGridColumn59.Header.VisiblePosition = 6;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn53,
																											ultraGridColumn54,
																											ultraGridColumn55,
																											ultraGridColumn56,
																											ultraGridColumn57,
																											ultraGridColumn58,
																											ultraGridColumn59});
			this.cmbArticuloSRI.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbArticuloSRI.DisplayMember = "Detalle";
			this.cmbArticuloSRI.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbArticuloSRI.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbArticuloSRI.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticuloSRI.Location = new System.Drawing.Point(392, 32);
			this.cmbArticuloSRI.Name = "cmbArticuloSRI";
			this.cmbArticuloSRI.Size = new System.Drawing.Size(328, 21);
			this.cmbArticuloSRI.TabIndex = 4;
			this.cmbArticuloSRI.ValueMember = "idArticuloSRI";
			this.cmbArticuloSRI.MouseDown += new System.Windows.Forms.MouseEventHandler(this.cmbArticuloSRI_MouseDown);
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(728, 28);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(40, 26);
			this.btGenerar.TabIndex = 5;
			this.btGenerar.Text = "SRI";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// ultraCombo1
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraCombo1.Appearance = appearance15;
			this.ultraCombo1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas7"));
			this.ultraCombo1.DataSource = this.cdvCuentaNombre;
			ultraGridColumn60.Header.VisiblePosition = 1;
			ultraGridColumn61.Header.VisiblePosition = 4;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 3;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 2;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 0;
			ultraGridColumn64.Width = 293;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn60,
																											ultraGridColumn61,
																											ultraGridColumn62,
																											ultraGridColumn63,
																											ultraGridColumn64});
			this.ultraCombo1.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.ultraCombo1.DisplayMember = "Descripcion";
			this.ultraCombo1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo1.Location = new System.Drawing.Point(312, 520);
			this.ultraCombo1.Name = "ultraCombo1";
			this.ultraCombo1.Size = new System.Drawing.Size(200, 21);
			this.ultraCombo1.TabIndex = 17;
			this.ultraCombo1.ValueMember = "idCuenta";
			this.ultraCombo1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.ultraCombo1_MouseUp);
			// 
			// cmbEgreso
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEgreso.Appearance = appearance16;
			this.cmbEgreso.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEgreso.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEgreso.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas7"));
			this.cmbEgreso.DataSource = this.cdvCuenta;
			ultraGridColumn65.Header.VisiblePosition = 2;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 4;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 3;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 0;
			ultraGridColumn69.Header.VisiblePosition = 1;
			ultraGridColumn69.Width = 239;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn65,
																											ultraGridColumn66,
																											ultraGridColumn67,
																											ultraGridColumn68,
																											ultraGridColumn69});
			this.cmbEgreso.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbEgreso.DisplayMember = "CodRapido";
			this.cmbEgreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEgreso.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEgreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEgreso.Location = new System.Drawing.Point(96, 520);
			this.cmbEgreso.Name = "cmbEgreso";
			this.cmbEgreso.Size = new System.Drawing.Size(200, 21);
			this.cmbEgreso.TabIndex = 16;
			this.cmbEgreso.ValueMember = "idCuenta";
			this.cmbEgreso.MouseUp += new System.Windows.Forms.MouseEventHandler(this.ultraCombo1_MouseUp);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(16, 520);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(52, 16);
			this.label5.TabIndex = 33;
			this.label5.Text = "Egreso B.";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.label5, "Egreso de Bodega");
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(16, 544);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(72, 16);
			this.label7.TabIndex = 76;
			this.label7.Text = "Transferencia";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.label7, "Egreso de Bodega");
			// 
			// cmbTransfCod
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTransfCod.Appearance = appearance17;
			this.cmbTransfCod.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTransfCod.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTransfCod.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas8"));
			this.cmbTransfCod.DataSource = this.cdvCuentaNombre;
			ultraGridColumn70.Header.VisiblePosition = 1;
			ultraGridColumn71.Header.VisiblePosition = 4;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 3;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 2;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 0;
			ultraGridColumn74.Width = 293;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn70,
																											ultraGridColumn71,
																											ultraGridColumn72,
																											ultraGridColumn73,
																											ultraGridColumn74});
			this.cmbTransfCod.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.cmbTransfCod.DisplayMember = "Descripcion";
			this.cmbTransfCod.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTransfCod.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTransfCod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTransfCod.Location = new System.Drawing.Point(312, 544);
			this.cmbTransfCod.Name = "cmbTransfCod";
			this.cmbTransfCod.Size = new System.Drawing.Size(200, 21);
			this.cmbTransfCod.TabIndex = 19;
			this.cmbTransfCod.ValueMember = "idCuenta";
			this.cmbTransfCod.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbTransfCod_MouseUp);
			// 
			// cmbTransf
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTransf.Appearance = appearance18;
			this.cmbTransf.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTransf.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTransf.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGrupoArt, "ArticuloGrupo.idCuentas8"));
			this.cmbTransf.DataSource = this.cdvCuenta;
			ultraGridColumn75.Header.VisiblePosition = 2;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 4;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 3;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 0;
			ultraGridColumn79.Header.VisiblePosition = 1;
			ultraGridColumn79.Width = 239;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn75,
																											ultraGridColumn76,
																											ultraGridColumn77,
																											ultraGridColumn78,
																											ultraGridColumn79});
			this.cmbTransf.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbTransf.DisplayMember = "CodRapido";
			this.cmbTransf.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTransf.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTransf.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTransf.Location = new System.Drawing.Point(96, 544);
			this.cmbTransf.Name = "cmbTransf";
			this.cmbTransf.Size = new System.Drawing.Size(200, 21);
			this.cmbTransf.TabIndex = 18;
			this.cmbTransf.ValueMember = "idCuenta";
			this.cmbTransf.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbTransfCod_MouseUp);
			// 
			// btAuditoria
			// 
			this.btAuditoria.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance19.Image = ((object)(resources.GetObject("appearance19.Image")));
			this.btAuditoria.Appearance = appearance19;
			this.btAuditoria.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAuditoria.Location = new System.Drawing.Point(376, 664);
			this.btAuditoria.Name = "btAuditoria";
			this.btAuditoria.Size = new System.Drawing.Size(24, 26);
			this.btAuditoria.TabIndex = 124;
			this.btAuditoria.Click += new System.EventHandler(this.btAuditoria_Click);
			this.btAuditoria.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btAuditoria_MouseUp);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(528, 424);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(76, 16);
			this.label8.TabIndex = 126;
			this.label8.Text = "Distribuidor PI";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label8.Visible = false;
			// 
			// spnIncr1
			// 
			this.spnIncr1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvGrupoArt, "Incr1"));
			this.spnIncr1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnIncr1.Enabled = false;
			this.spnIncr1.FormatString = "#,##0.00";
			this.spnIncr1.Location = new System.Drawing.Point(616, 424);
			this.spnIncr1.MaxValue = 100;
			this.spnIncr1.MinValue = 0;
			this.spnIncr1.Name = "spnIncr1";
			this.spnIncr1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnIncr1.PromptChar = ' ';
			this.spnIncr1.Size = new System.Drawing.Size(128, 22);
			this.spnIncr1.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnIncr1.TabIndex = 20;
			this.spnIncr1.Visible = false;
			this.spnIncr1.Click += new System.EventHandler(this.spnIncr1_Click);
			this.spnIncr1.ValueChanged += new System.EventHandler(this.spnIncr1_ValueChanged);
			this.spnIncr1.Enter += new System.EventHandler(this.spnIncr1_Enter);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(528, 448);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(62, 16);
			this.label9.TabIndex = 128;
			this.label9.Text = "Contado PA";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label9.Visible = false;
			// 
			// spnIncr2
			// 
			this.spnIncr2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvGrupoArt, "Incr2"));
			this.spnIncr2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnIncr2.Enabled = false;
			this.spnIncr2.FormatString = "#,##0.00";
			this.spnIncr2.Location = new System.Drawing.Point(616, 448);
			this.spnIncr2.MaxValue = 100;
			this.spnIncr2.MinValue = 0;
			this.spnIncr2.Name = "spnIncr2";
			this.spnIncr2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnIncr2.PromptChar = ' ';
			this.spnIncr2.Size = new System.Drawing.Size(128, 22);
			this.spnIncr2.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnIncr2.TabIndex = 21;
			this.spnIncr2.Visible = false;
			this.spnIncr2.Click += new System.EventHandler(this.spnIncr2_Click);
			this.spnIncr2.Enter += new System.EventHandler(this.spnIncr2_Enter);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(528, 472);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(56, 16);
			this.label10.TabIndex = 130;
			this.label10.Text = "Tarjeta PA";
			this.label10.Visible = false;
			// 
			// spnIncr3
			// 
			this.spnIncr3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvGrupoArt, "Incr3"));
			this.spnIncr3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnIncr3.Enabled = false;
			this.spnIncr3.FormatString = "#,##0.00";
			this.spnIncr3.Location = new System.Drawing.Point(616, 472);
			this.spnIncr3.MaxValue = 100;
			this.spnIncr3.MinValue = 0;
			this.spnIncr3.Name = "spnIncr3";
			this.spnIncr3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnIncr3.PromptChar = ' ';
			this.spnIncr3.Size = new System.Drawing.Size(72, 22);
			this.spnIncr3.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnIncr3.TabIndex = 129;
			this.spnIncr3.Visible = false;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(528, 496);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(63, 16);
			this.label11.TabIndex = 132;
			this.label11.Text = "Crédito PVP";
			this.label11.Visible = false;
			// 
			// spnIncr4
			// 
			this.spnIncr4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvGrupoArt, "Incr4"));
			this.spnIncr4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnIncr4.Enabled = false;
			this.spnIncr4.FormatString = "#,##0.00";
			this.spnIncr4.Location = new System.Drawing.Point(616, 496);
			this.spnIncr4.MaxValue = 100;
			this.spnIncr4.MinValue = 0;
			this.spnIncr4.Name = "spnIncr4";
			this.spnIncr4.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnIncr4.PromptChar = ' ';
			this.spnIncr4.Size = new System.Drawing.Size(72, 22);
			this.spnIncr4.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnIncr4.TabIndex = 131;
			this.spnIncr4.Visible = false;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(528, 520);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(30, 16);
			this.label12.TabIndex = 134;
			this.label12.Text = "Incr5";
			this.label12.Visible = false;
			// 
			// spnIncr5
			// 
			this.spnIncr5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvGrupoArt, "Incr5"));
			this.spnIncr5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnIncr5.Enabled = false;
			this.spnIncr5.FormatString = "#,##0.00";
			this.spnIncr5.Location = new System.Drawing.Point(616, 520);
			this.spnIncr5.MaxValue = 100;
			this.spnIncr5.MinValue = 0;
			this.spnIncr5.Name = "spnIncr5";
			this.spnIncr5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnIncr5.PromptChar = ' ';
			this.spnIncr5.Size = new System.Drawing.Size(72, 22);
			this.spnIncr5.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnIncr5.TabIndex = 133;
			this.spnIncr5.Visible = false;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(528, 544);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(30, 16);
			this.label13.TabIndex = 136;
			this.label13.Text = "Incr6";
			this.label13.Visible = false;
			// 
			// spnIncr6
			// 
			this.spnIncr6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdvGrupoArt, "Incr6"));
			this.spnIncr6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnIncr6.Enabled = false;
			this.spnIncr6.FormatString = "#,##0.00";
			this.spnIncr6.Location = new System.Drawing.Point(616, 544);
			this.spnIncr6.MaxValue = 100;
			this.spnIncr6.MinValue = 0;
			this.spnIncr6.Name = "spnIncr6";
			this.spnIncr6.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnIncr6.PromptChar = ' ';
			this.spnIncr6.Size = new System.Drawing.Size(72, 22);
			this.spnIncr6.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnIncr6.TabIndex = 135;
			this.spnIncr6.TabStop = false;
			this.spnIncr6.Visible = false;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.White;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label14.Location = new System.Drawing.Point(552, 400);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(185, 17);
			this.label14.TabIndex = 137;
			this.label14.Text = "% DE INCREMENTO DE PRECIO";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label14.Visible = false;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(752, 424);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(15, 16);
			this.label15.TabIndex = 138;
			this.label15.Text = "%";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label15.Visible = false;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(752, 448);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(15, 16);
			this.label16.TabIndex = 139;
			this.label16.Text = "%";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label16.Visible = false;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(696, 480);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(15, 16);
			this.label17.TabIndex = 140;
			this.label17.Text = "%";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label17.Visible = false;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(696, 504);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(15, 16);
			this.label18.TabIndex = 141;
			this.label18.Text = "%";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label18.Visible = false;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(696, 528);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(15, 16);
			this.label19.TabIndex = 142;
			this.label19.Text = "%";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label19.Visible = false;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(696, 552);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(15, 16);
			this.label20.TabIndex = 143;
			this.label20.Text = "%";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label20.Visible = false;
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
			// btnIncr
			// 
			this.btnIncr.Location = new System.Drawing.Point(608, 584);
			this.btnIncr.Name = "btnIncr";
			this.btnIncr.Size = new System.Drawing.Size(152, 32);
			this.btnIncr.TabIndex = 144;
			this.btnIncr.Text = "Cambiar % de Incremento ";
			this.btnIncr.Visible = false;
			this.btnIncr.Click += new System.EventHandler(this.btnIncr_Click);
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(312, 32);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(22, 16);
			this.label21.TabIndex = 145;
			this.label21.Text = "SRI";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance20.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance20;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand17.AddButtonCaption = "CompraFormaPago";
			ultraGridColumn80.Header.VisiblePosition = 0;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn80.Width = 93;
			ultraGridColumn81.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn81.Header.Caption = "Subgrupo";
			ultraGridColumn81.Header.VisiblePosition = 1;
			ultraGridColumn81.Width = 227;
			ultraGridColumn82.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn82.CellAppearance = appearance21;
			ultraGridColumn82.Format = "#,##0.00";
			ultraGridColumn82.Header.Caption = "% PI";
			ultraGridColumn82.Header.VisiblePosition = 2;
			ultraGridColumn82.Width = 56;
			ultraGridColumn83.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn83.CellAppearance = appearance22;
			ultraGridColumn83.Format = "#,##0.00";
			ultraGridColumn83.Header.Caption = "% PA";
			ultraGridColumn83.Header.VisiblePosition = 3;
			ultraGridColumn83.Width = 59;
			ultraGridColumn84.Header.VisiblePosition = 4;
			ultraGridColumn84.Width = 45;
			ultraGridBand17.Columns.AddRange(new object[] {
																											ultraGridColumn80,
																											ultraGridColumn81,
																											ultraGridColumn82,
																											ultraGridColumn83,
																											ultraGridColumn84});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance25.FontData.BoldAsString = "True";
			appearance25.FontData.Name = "Arial";
			appearance25.FontData.SizeInPoints = 10F;
			appearance25.ForeColor = System.Drawing.Color.White;
			appearance25.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance27;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(9, 80);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(408, 312);
			this.ultraGrid1.TabIndex = 146;
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			this.ultraGrid1.Click += new System.EventHandler(this.ultraGrid1_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(System.Double);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(System.Double);
			ultraDataColumn4.DefaultValue = 0;
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn5.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// ultraGrid3
			// 
			this.ultraGrid3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid3.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid3.DataSource = this.ultraDataSource3;
			appearance28.BackColor = System.Drawing.Color.White;
			this.ultraGrid3.DisplayLayout.Appearance = appearance28;
			this.ultraGrid3.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn85.Header.VisiblePosition = 0;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn85.Width = 170;
			ultraGridColumn86.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn86.Header.VisiblePosition = 1;
			ultraGridColumn86.Width = 191;
			ultraGridColumn87.Header.VisiblePosition = 2;
			ultraGridColumn87.Width = 36;
			ultraGridBand18.Columns.AddRange(new object[] {
																											ultraGridColumn85,
																											ultraGridColumn86,
																											ultraGridColumn87});
			this.ultraGrid3.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid3.DisplayLayout.Override.ActiveRowAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid3.DisplayLayout.Override.CardAreaAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance31.FontData.BoldAsString = "True";
			appearance31.FontData.Name = "Arial";
			appearance31.FontData.SizeInPoints = 10F;
			appearance31.ForeColor = System.Drawing.Color.White;
			appearance31.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid3.DisplayLayout.Override.HeaderAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.RowSelectorAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.SelectedRowAppearance = appearance33;
			this.ultraGrid3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid3.Location = new System.Drawing.Point(681, 80);
			this.ultraGrid3.Name = "ultraGrid3";
			this.ultraGrid3.Size = new System.Drawing.Size(248, 312);
			this.ultraGrid3.TabIndex = 152;
			this.ultraGrid3.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid3_CellChange);
			this.ultraGrid3.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid3_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn8.DefaultValue = false;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance34.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance34;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn88.Header.VisiblePosition = 0;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn88.Width = 97;
			ultraGridColumn89.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn89.Header.VisiblePosition = 1;
			ultraGridColumn89.Width = 188;
			ultraGridColumn90.Header.VisiblePosition = 2;
			ultraGridColumn90.Width = 39;
			ultraGridBand19.Columns.AddRange(new object[] {
																											ultraGridColumn88,
																											ultraGridColumn89,
																											ultraGridColumn90});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand19);
			appearance35.ForeColor = System.Drawing.Color.Black;
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance37.FontData.BoldAsString = "True";
			appearance37.FontData.Name = "Arial";
			appearance37.FontData.SizeInPoints = 10F;
			appearance37.ForeColor = System.Drawing.Color.White;
			appearance37.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance39;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(425, 80);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(248, 312);
			this.ultraGrid2.TabIndex = 151;
			this.ultraGrid2.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_CellChange);
			this.ultraGrid2.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_AfterCellUpdate);
			this.ultraGrid2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid2_InitializeLayout);
			this.ultraGrid2.Enter += new System.EventHandler(this.ultraGrid2_Enter);
			this.ultraGrid2.Click += new System.EventHandler(this.ultraGrid2_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn11.DataType = typeof(bool);
			ultraDataColumn11.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// btSubGrupo
			// 
			this.btSubGrupo.Location = new System.Drawing.Point(8, 56);
			this.btSubGrupo.Name = "btSubGrupo";
			this.btSubGrupo.Size = new System.Drawing.Size(88, 24);
			this.btSubGrupo.TabIndex = 153;
			this.btSubGrupo.Text = "Subgrupos";
			this.btSubGrupo.Click += new System.EventHandler(this.btSubGrupo_Click);
			// 
			// btCategorias
			// 
			this.btCategorias.Location = new System.Drawing.Point(424, 56);
			this.btCategorias.Name = "btCategorias";
			this.btCategorias.Size = new System.Drawing.Size(80, 24);
			this.btCategorias.TabIndex = 154;
			this.btCategorias.Text = "Categorias";
			this.btCategorias.Click += new System.EventHandler(this.btCategorias_Click);
			// 
			// btMarcas
			// 
			this.btMarcas.Location = new System.Drawing.Point(680, 56);
			this.btMarcas.Name = "btMarcas";
			this.btMarcas.Size = new System.Drawing.Size(72, 24);
			this.btMarcas.TabIndex = 155;
			this.btMarcas.Text = "Marcas";
			this.btMarcas.Click += new System.EventHandler(this.btMarcas_Click);
			// 
			// txtBuscaSubGrupo
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscaSubGrupo.Appearance = appearance40;
			this.txtBuscaSubGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscaSubGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscaSubGrupo.Enabled = false;
			this.txtBuscaSubGrupo.Location = new System.Drawing.Point(104, 56);
			this.txtBuscaSubGrupo.Name = "txtBuscaSubGrupo";
			this.txtBuscaSubGrupo.Size = new System.Drawing.Size(312, 22);
			this.txtBuscaSubGrupo.TabIndex = 156;
			this.txtBuscaSubGrupo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscaSubGrupo_KeyPress);
			this.txtBuscaSubGrupo.ValueChanged += new System.EventHandler(this.txtBuscaSubGrupo_ValueChanged);
			// 
			// txtBuscarCategoria
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscarCategoria.Appearance = appearance41;
			this.txtBuscarCategoria.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscarCategoria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscarCategoria.Enabled = false;
			this.txtBuscarCategoria.Location = new System.Drawing.Point(512, 56);
			this.txtBuscarCategoria.Name = "txtBuscarCategoria";
			this.txtBuscarCategoria.Size = new System.Drawing.Size(160, 22);
			this.txtBuscarCategoria.TabIndex = 157;
			this.txtBuscarCategoria.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscarCategoria_KeyPress);
			this.txtBuscarCategoria.ValueChanged += new System.EventHandler(this.txtBuscarCategoria_ValueChanged);
			// 
			// txtBuscarMarca
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscarMarca.Appearance = appearance42;
			this.txtBuscarMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscarMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscarMarca.Enabled = false;
			this.txtBuscarMarca.Location = new System.Drawing.Point(760, 56);
			this.txtBuscarMarca.Name = "txtBuscarMarca";
			this.txtBuscarMarca.Size = new System.Drawing.Size(168, 22);
			this.txtBuscarMarca.TabIndex = 158;
			this.txtBuscarMarca.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscarMarca_KeyPress);
			this.txtBuscarMarca.ValueChanged += new System.EventHandler(this.txtBuscarMarca_ValueChanged);
			// 
			// ArtGrupo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(938, 624);
			this.Controls.Add(this.txtBuscarMarca);
			this.Controls.Add(this.txtBuscarCategoria);
			this.Controls.Add(this.txtBuscaSubGrupo);
			this.Controls.Add(this.btMarcas);
			this.Controls.Add(this.btCategorias);
			this.Controls.Add(this.btSubGrupo);
			this.Controls.Add(this.ultraGrid3);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblCosto);
			this.Controls.Add(this.lblVenta);
			this.Controls.Add(this.lblCompra);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnIncr);
			this.Controls.Add(this.spnIncr6);
			this.Controls.Add(this.spnIncr5);
			this.Controls.Add(this.spnIncr4);
			this.Controls.Add(this.spnIncr3);
			this.Controls.Add(this.spnIncr2);
			this.Controls.Add(this.spnIncr1);
			this.Controls.Add(this.btAuditoria);
			this.Controls.Add(this.cmbTransfCod);
			this.Controls.Add(this.cmbTransf);
			this.Controls.Add(this.ultraCombo1);
			this.Controls.Add(this.cmbEgreso);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.cmbArticuloSRI);
			this.Controls.Add(this.cmbDevolucionD);
			this.Controls.Add(this.cmbDevolucionC);
			this.Controls.Add(this.cmbDescD);
			this.Controls.Add(this.cmbDescC);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbCostoD);
			this.Controls.Add(this.cmbVentaD);
			this.Controls.Add(this.cmbCompraD);
			this.Controls.Add(this.cmbCostoC);
			this.Controls.Add(this.cmbVentaC);
			this.Controls.Add(this.cmbCompraC);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtIdGrupoArt);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ArtGrupo";
			this.ShowInTaskbar = false;
			this.Text = "Grupos de Artículos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ArtGrupo_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ArtGrupo_Closing);
			this.Load += new System.EventHandler(this.ArtGrupo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuentaNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCompraC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVentaC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCostoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCostoD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVentaD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCompraD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDescD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDescC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDevolucionD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDevolucionC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticuloSRI)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEgreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransfCod)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransf)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnIncr6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscaSubGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscarCategoria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscarMarca)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			#region Verificaciones
			if (this.txtIdGrupoArt.Value == null) return;
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar Grupos de Items",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (MessageBox.Show("¿Desea Borrar este Grupo de Artículos?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
				MessageBoxDefaultButton.Button2) == DialogResult.No) 
				return;
			#endregion Verificaciones

			// Verificacion de Existencia de articulos en compras
			string stSelect = "SELECT COUNT(*) FROM Articulo WHERE Articulo.idGrupoArticulo = "
				+ this.txtIdGrupoArt.Value.ToString().Trim();
			int iCuenta = (int) Funcion.iEscalarSQL(cdsGrupoArt, stSelect);
			if (iCuenta > 0) 
			{
				MessageBox.Show("No puede borrar este Grupo de artículos.\n"
					+"Tiene " + iCuenta.ToString().Trim() + " registros relacionados."
					+ " Eliminación CANCELADA",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			this.barraDato1.BorraRegistro();
		}

		private void BuscaGrupo(int id)
		{
			this.barraDato1.IdRegistro = id;
			this.barraDato1.ProximoId(5);
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (BuscaGrupoArt1 miBusqueda = new BuscaGrupoArt1())
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					try
					{
						int id = int.Parse(miBusqueda.idBusca.Text);
						BuscaGrupo(id);
					}
					catch{};
				}
			}
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a Crear Nuevos Grupos de Items", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
//			spnIncr1.ReadOnly  = true;
//			spnIncr2.ReadOnly  = true;
//			spnIncr3.ReadOnly  = true;
//			spnIncr4.ReadOnly  = true;
//			spnIncr5.ReadOnly  = true;
//			spnIncr6.ReadOnly  = true;
//			string stActivado = "Select IsNull(Activado, 0) From SeteoG Where Codigo = 'FGARTPOINT'";
//			if (Funcion.bEjecutaSQL(cdsSeteoF, stActivado))
//			{
//				//btnIncr.Enabled = true;
//				spnIncr1.ReadOnly	 = false;
//				spnIncr2.ReadOnly  = false;
//				spnIncr3.ReadOnly  = false;
//				spnIncr4.ReadOnly  = false;
//				spnIncr5.ReadOnly  = false;
//				spnIncr6.ReadOnly  = false;
//			}	
			
			this.barraDato1.CrearRegistro();

			this.txtBuscaSubGrupo.Enabled = true;
			this.txtBuscarMarca.Enabled = true;
			this.txtBuscarCategoria.Enabled = true;
			this.txtBuscaSubGrupo.Text = "";
			this.txtBuscarCategoria.Text = "";
			this.txtBuscarMarca.Text = "";

			this.cmbArticuloSRI.Enabled = false;
			
			CuentaDefault();

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaSubGrupos 0, false, ''");
//			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCategorias 0, false");
			//this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaMarcas 0, false");

			this.txtNombre.Focus();
		}

		private void CuentaDefault()
		{
			string sSQL = string.Format("Exec CuentaTipoGrupo {0}", (int)this.cmbTipo.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				drGrupo["idArticuloSRI"] = (int)dr.GetValue(0);
				drGrupo["idCuentas2"] = (int)dr.GetValue(1);
				drGrupo["idCuentas3"] = (int)dr.GetValue(2);
				drGrupo["idCuentas4"] = (int)dr.GetValue(3);
				drGrupo["idCuentas5"] = (int)dr.GetValue(4);
				drGrupo["idCuentas6"] = (int)dr.GetValue(5);
				drGrupo["idCuentas7"] = (int)dr.GetValue(6);
				drGrupo["idCuentas8"] = (int)dr.GetValue(7);
				dr.Close();
			}
			else
			{
				dr.Close();

				drGrupo["idArticuloSRI"] = 0;// System.DBNull.Value;
				drGrupo["idCuentas2"] = 0;//System.DBNull.Value;
				drGrupo["idCuentas3"] = 0;// System.DBNull.Value;
				drGrupo["idCuentas4"] = 0;// System.DBNull.Value;
				drGrupo["idCuentas5"] = 0;// System.DBNull.Value;
				drGrupo["idCuentas6"] = 0;// System.DBNull.Value;
				drGrupo["idCuentas7"] = 0;// System.DBNull.Value;
				drGrupo["idCuentas8"] = 0;// System.DBNull.Value;

				MessageBox.Show(string.Format("No ha predeterminado cuentas para el tipo '{0}'.", this.cmbTipo.Text.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}	
			dr.Close();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.BorraErrores();
			this.barraDato1.DeshacerRegistro();

			this.txtBuscaSubGrupo.Enabled = false;
			this.txtBuscarMarca.Enabled = false;
			this.txtBuscarCategoria.Enabled = false;
			this.txtBuscaSubGrupo.Text = "";
			this.txtBuscarCategoria.Text = "";
			this.txtBuscarMarca.Text = "";

			//btnIncr.Enabled = false;
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (this.txtIdGrupoArt.Value == null) return;

			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar Grupos de Items", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
//			spnIncr1.ReadOnly  = true;
//			spnIncr2.ReadOnly  = true;
//			spnIncr3.ReadOnly  = true;
//			spnIncr4.ReadOnly  = true;
//			spnIncr5.ReadOnly  = true;
//			spnIncr6.ReadOnly  = true;
//
//			string stActivado = "Select IsNull(Activado, 0) From SeteoG Where Codigo = 'FGARTPOINT'";
//			if (Funcion.bEjecutaSQL(cdsSeteoF, stActivado))
//			{
//				//btnIncr.Enabled = true;
//				spnIncr1.ReadOnly	 = false;
//				spnIncr2.ReadOnly  = false;
//				spnIncr3.ReadOnly  = false;
//				spnIncr4.ReadOnly  = false;
//				spnIncr5.ReadOnly  = false;
//				spnIncr6.ReadOnly  = false;
//			}	
			this.barraDato1.EditarRegistro();

			this.txtBuscaSubGrupo.Enabled = true;
			this.txtBuscarMarca.Enabled = true;
			this.txtBuscarCategoria.Enabled = true;
			this.txtBuscaSubGrupo.Text = "";
			this.txtBuscarCategoria.Text = "";
			this.txtBuscarMarca.Text = "";

			cmbTipo_ValueChanged(this, e);
			string sSQL = string.Format("Exec ListaSubGrupos {0}, true, '{1}'", (int)drGrupo["idGrupoArticulo"], this.txtBuscaSubGrupo.Text.ToString());
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void BorraErrores()
		{
			this.errorProvider.SetError(this.txtNombre, "");
			this.errorProvider.SetError(this.cmbTipo, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			this.BorraErrores();
			string stGrupo = (string) this.txtNombre.Value;
			if (txtNombre.Value == System.DBNull.Value || txtNombre.Value == null)
			{
				this.errorProvider.SetError(this.txtNombre, "Ingrese Grupo");
				bOk = false;
			}
			else
			{
				// Verificacion de Codigo Repetido
				this.cdsGrupoArt.Schema.Connections[0].Open();
				C1.Data.SchemaObjects.Connection miConeccion = this.cdsGrupoArt.Schema.Connections[0];
				string stSelect = "SELECT COUNT(*) FROM ArticuloGrupo WHERE Grupo = '" +
					this.txtNombre.Value.ToString().Trim() + "'";

				if (!this.barraDato1.bNuevo)
					stSelect += " And idGrupoArticulo <> " +
						this.txtIdGrupoArt.Value.ToString().Trim();

				SqlCommand cmdCodigo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				int iCuenta = (int) cmdCodigo.ExecuteScalar();
				if (iCuenta > 0) 
				{
					this.errorProvider.SetError(this.txtNombre,"Grupo de Artículo ya existe");
					bOk = false;
				}
				this.cdsGrupoArt.Schema.Connections[0].Close();
				// Fin de verificacion de Código repetido
			}

			if ((int) this.cmbTipo.Value < 1)
			{
				this.errorProvider.SetError(this.cmbTipo, "Ingrese Tipo de Grupo");
				bOk = false;
			}
			return bOk;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;

			#region Verifica porcentajes de incremento de precio mayores a cero
//			if ((int)this.cmbTipo.Value == 1)
//			{
//				if ((double)this.spnIncr1.Value == 0)
//				{
//					MessageBox.Show("El porcentaje de incremento para precio distribuidor no pueder ser igual a cero.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
//					this.spnIncr1.Focus();
//					return;
//				}
//
//				if ((double)this.spnIncr2.Value == 0)
//				{
//					MessageBox.Show("El porcentaje de incremento para precio contado no pueder ser igual a cero.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
//					this.spnIncr2.Focus();
//					return;
//				}
//			}
			#endregion Verifica porcentajes de incremento de precio mayores a cero

			this.barraDato1.GrabaRegistro();
			// Cambia el tipo de todos los Articulos del grupo
			SqlCommand cmdActualiza;
			this.cdsArticulo.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = this.cdsArticulo.Schema.Connections[0];
			cmdActualiza = new SqlCommand ("UPDATE Articulo SET idTipoGrupo = " + this.cmbTipo.Value.ToString() + " WHERE idGrupoArticulo = " + this.txtIdGrupoArt.Text, (SqlConnection) miConeccion.DbConnection);

			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			cmdActualiza.ExecuteNonQuery();
			this.cdsArticulo.Schema.Connections[0].Close();

			int idGrupo = (int)drGrupo["idGrupoArticulo"];
			string sSQLGrupoSubgrupo = "";
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				int idSubgrupo = (int)dr.Cells["idSubGrupo"].Value;
				bool Sel = (bool)dr.Cells["Sel"].Value;
				sSQLGrupoSubgrupo = string.Format("Exec ActualizaGrupoSubgrupo {0}, {1}, {2}", idGrupo, idSubgrupo, Sel);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLGrupoSubgrupo, true);
			}

			#region Actualiza el porcentaje de incremento de precio
//			string stExec = "Exec Point_IncrementoPrecio  " + this.txtIdGrupoArt.Value.ToString().Trim();
//			Funcion.EjecutaSQL(cdsArticulo, stExec, true);
			#endregion Actualiza el porcentaje de incremento de precio

			this.txtBuscaSubGrupo.Enabled = false;
			this.txtBuscarMarca.Enabled = false;
			this.txtBuscarCategoria.Enabled = false;
			this.txtBuscaSubGrupo.Text = "";
			this.txtBuscarCategoria.Text = "";
			this.txtBuscarMarca.Text = "";
		}

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}
    
		private void ArtGrupo_Load(object sender, System.EventArgs e)
		{
//			string stActivado = "Select IsNull(Activado, 0) From SeteoG Where Codigo = 'FGARTPOINT'";
//			if (Funcion.bEjecutaSQL(cdsSeteoF, stActivado))
//			{
//				this.Width = 720;
//			}
//			if (!Funcion.Permiso("77",cdsGrupoArt))
//			{
//				MessageBox.Show("No puede ingresar a Grupos de Artículos", "Informacion");
//				BeginInvoke(new MethodInvoker(UnloadMe));
//				return;
//			}

//			string sFiltro = "";

//			if (IInventario != 0)
//			{
//				if (IInventario == 1)
//				{
//					barraDato1.stFiltroExt = " And idTipoGrupo In (1, 5, 6, 13)";  
//					sFiltro = "idTipoGrupo In (1, 5, 6, 13)";
//				}
//				else if (IInventario == 2)
//				{
//					barraDato1.stFiltroExt = " And idTipoGrupo Not In (1, 5, 6, 13, 14, 15)"; 
//					sFiltro = "idTipoGrupo Not In (1, 5, 6, 13, 14, 15)";
//				}
//
//				cdvTipo.RowFilter = sFiltro;
//			}

			Cursor = Cursors.Default;
			this.txtIdGrupoArt.Width = 0;
			if (IdGrupo > 0) this.BuscaGrupo(IdGrupo);
			else this.barraDato1.PosUltima();

			this.barraDato1.HabilitaControles(false);
			string stAudita = "Exec AuditaCrear 53, 6, 'Art. Grupo'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			miAcceso = new Acceso(cdsArticulo, "0410A");
		}

		private void cdsGrupoArt_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupoArt.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirConta);				
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			drGrupo = cdsGrupoArt.TableViews["ArticuloGrupo"].Rows[0];

			string sSQL = string.Format("Exec ListaSubGrupos {0}, true, '{1}'", (int)drGrupo["idGrupoArticulo"], this.txtBuscaSubGrupo.Text.ToString());
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			
			string sSQL1 = string.Format("Declare @Sel Bit Set @Sel = 0 Select idCategoria, Categoria, @Sel As Sel From ArticuloCategoria Where idCategoria = 0");//Exec ListaCategorias 0, false");
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL1);

			sSQL = string.Format("Declare @Sel Bit Set @Sel = 0 Select idMarca, Marca, @Sel As Sel From ArticuloMarca Where idMarca = 0");// Exec ListaMarcas 0, false");
			this.ultraGrid3.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			if (cmbTipo.Value != null && cmbTipo.Value != System.DBNull.Value && (int) cmbTipo.Value == 5)
			{
				lblCosto.Text = "Depreciac.";
				lblVenta.Text = "Dep. Acum.";
			}
			else
			{
				lblCosto.Text = "Costo";
				lblVenta.Text = "Venta";
			}
		}

		private void cmbTipo_ValueChanged(object sender, System.EventArgs e)
		{
			if (cmbTipo.Value == System.DBNull.Value || !barraDato1.bEditar) return;
			barraDatoSQL1_Refresca(this, e);
			cmbCompraC.Enabled = true;
			cmbCompraD.Enabled = true;
			cmbVentaC.Enabled = true;
			cmbVentaD.Enabled = true;
			cmbCostoC.Enabled = true;
			cmbCostoD.Enabled = true;


			switch ((int) cmbTipo.Value)
			{
				case 1:
					//this.cmbArticuloSRI.Value = 53;
					this.cmbArticuloSRI.Enabled = false;
					break;
				case 2: // Gasto
					this.cmbArticuloSRI.Value = DBNull.Value;
					this.cmbArticuloSRI.Enabled = true;
					cmbCostoC.Enabled = false;
					cmbCostoD.Enabled = false;
					cmbVentaC.Enabled = false;
					cmbVentaD.Enabled = false;
					break;
				case 3: // Servicio
					this.cmbArticuloSRI.Value = DBNull.Value;
					this.cmbArticuloSRI.Enabled = true;
					cmbCostoC.Enabled = false;
					cmbCostoD.Enabled = false;
					break;
				case 4: // Saldo Incial
					this.cmbArticuloSRI.Value = DBNull.Value;
					this.cmbArticuloSRI.Enabled = true;
					cmbCompraC.Enabled = false;
					cmbCompraD.Enabled = false;
					cmbCostoC.Enabled = false;
					cmbCostoD.Enabled = false;
					cmbVentaC.Enabled = false;
					cmbVentaD.Enabled = false;
					break;
				case 5:
				case 6:
				case 7:
				case 8:
				case 9:
				case 10:
				case 11:
				case 12:
					this.cmbArticuloSRI.Value = DBNull.Value;
					this.cmbArticuloSRI.Enabled = true;
					break;
			}	
	
			CuentaDefault();
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			if (cmbArticuloSRI.Value == System.DBNull.Value) return;
			if (barraDato1.bEditar)
			{
				MessageBox.Show("No puede asignar articulos mientras edita el Grupo de Articulos", "Información");
				return;
			}

			int IdArtSRI = (int) cmbArticuloSRI.Value;
			string stCod = cmbArticuloSRI.Text;
			if (DialogResult.No == MessageBox.Show("¿Desea Asignar el codigo "
				+ stCod + " a todos los articulos del grupo?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			string stSelect = "Update Articulo Set idArticuloSRI = " 
				+ IdArtSRI.ToString() + " WHERE idGrupoArticulo = " 
				+	this.txtIdGrupoArt.Text;
			Funcion.EjecutaSQL(cdsArticulo, stSelect, true);
			MessageBox.Show("Articulos Actualizados", "Información");
		}

		private void cmbCompraD_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbCompraD.Value = 0;
			cmbCompraD.Focus();
			cmbCompraD.Refresh();
		}

		private void cmbVentaD_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbVentaD.Value = 0;
			cmbVentaD.Focus();
			cmbVentaD.Refresh();
		}

		private void cmbCostoD_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbCostoD.Value = 0;
			cmbCostoD.Focus();
			cmbCostoD.Refresh();
		}

		private void cmbDescD_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbDescD.Value = 0;
			cmbDescD.Focus();
			cmbDescD.Refresh();
		}

		private void cmbDevolucionD_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbDevolucionD.Value = 0;
			cmbDevolucionD.Focus();
			cmbDevolucionD.Refresh();
		}

		private void cmbArticuloSRI_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbArticuloSRI.Value = 0;
			cmbArticuloSRI.Focus();
			cmbArticuloSRI.Refresh();
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			ArticuloTipoIng miArt = new ArticuloTipoIng();
			miArt.Show();
		}

		private void ArtGrupo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Art. Grupo'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			this.DialogResult = DialogResult.OK;
		}

		private void btAuditoria_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(29, int.Parse(txtIdGrupoArt.Value.ToString()), true);
			Cursor = Cursors.Default;
		}

		private void btAuditoria_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			Funcion.AuditoriaImprime(29, int.Parse(txtIdGrupoArt.Value.ToString()), false);
			Cursor = Cursors.Default;		
		}

		private void ultraCombo1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbEgreso.Value = 0;
			cmbEgreso.Focus();
			cmbEgreso.Refresh();
		}

		private void cmbTransfCod_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			cmbTransfCod.Value = 0;
			cmbTransfCod.Focus();
			cmbTransfCod.Refresh();
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnIncr_Click(object sender, System.EventArgs e)
		{
			if (barraDato1.bEditar)
			{
				MessageBox.Show("No puede realizar incrementos mientras edita el Grupo",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
				string stMensaje = "Desea Realizar el incremento" ;
						
				if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación",
					MessageBoxButtons.YesNo, MessageBoxIcon.Question,
					MessageBoxDefaultButton.Button2)) return;

				try
				{
					string stExec = "Exec Point_IncrementoPrecio  " + this.txtIdGrupoArt.Value.ToString().Trim();
					Funcion.EjecutaSQL(cdsArticulo, stExec, true);
				}
				catch(System.Exception)
				{
						
					MessageBox.Show("Error al realizar Incremento", "Información");	
				}
				
				MessageBox.Show("Incremento realizado con Exito", "Información");	
			}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void cmbTipo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipo_RowSelected(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
		{
		
		}

		private void cmbCompraC_Enter(object sender, System.EventArgs e)
		{
			this.cmbCompraC.Show();
		}

		private void spnIncr1_Click(object sender, System.EventArgs e)
		{
			this.spnIncr1.SelectAll();
		}

		private void spnIncr1_Enter(object sender, System.EventArgs e)
		{
			this.spnIncr1.SelectAll();
		}

		private void spnIncr2_Click(object sender, System.EventArgs e)
		{
			this.spnIncr2.SelectAll();
		}

		private void spnIncr2_Enter(object sender, System.EventArgs e)
		{
			this.spnIncr2.SelectAll();
		}

		private void spnIncr1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbVentaC_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null) return;
			string sSQL = string.Format("Exec ListaCategorias {0}, true, '{1}'", (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value, this.txtBuscarCategoria.Text.ToString());
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			sSQL = string.Format("Exec ListaMarcas {0}, true, '{1}'", (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value, this.txtBuscarMarca.Text.ToString());
			this.ultraGrid3.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void ultraGrid2_Click(object sender, System.EventArgs e)
		{
			
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid2_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (this.ultraGrid2.ActiveRow.Cells["idCategoria"].Value == DBNull.Value) return;
			
			bool Sel = (bool)this.ultraGrid2.ActiveRow.Cells["Sel"].Value;

			int idSubGrupo = (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value;
			int idCategoria = (int)this.ultraGrid2.ActiveRow.Cells["idCategoria"].Value;

			string sSQL = string.Format("Exec ActualizaSubGrupoCategoria {0}, {1}", idSubGrupo, idCategoria);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
		}

		private void ultraGrid2_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void ultraGrid2_Enter(object sender, System.EventArgs e)
		{
			
		}

		private void ultraGrid3_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (this.ultraGrid3.ActiveRow.Cells["idMarca"].Value == DBNull.Value)return;
			
			bool Sel = (bool)this.ultraGrid3.ActiveRow.Cells["Sel"].Value;

			int idSubGrupo = (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value;
			int idMarca = (int)this.ultraGrid3.ActiveRow.Cells["idMarca"].Value;

			string sSQL = string.Format("Exec ActualizaSubGrupoMarca {0}, {1}", idSubGrupo, idMarca);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
		}

		private void btSubGrupo_Click(object sender, System.EventArgs e)
		{
			using (ArtSubGrupo miSubgrupo = new ArtSubGrupo())
			{
				if (DialogResult.OK == miSubgrupo.ShowDialog())
				{					
					bool bNuevo = false;
					int idGrupo = (int)drGrupo["idGrupoArticulo"];
					if (idGrupo > -1) bNuevo = true;
					else bNuevo = false;
					string sSQL = string.Format("Exec ListaSubGrupos {0}, {1}, '{2}'", idGrupo, bNuevo, this.txtBuscaSubGrupo.Text.ToString());
					this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);				
				}
			}
		}

		private void btCategorias_Click(object sender, System.EventArgs e)
		{
			using (ArticuloCategoria miCategoria = new ArticuloCategoria())
			{
				if (DialogResult.OK == miCategoria.ShowDialog())
				{					
					int idSubGrupo = (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value;
					
					string sSQL = string.Format("Exec ListaCategorias {0}, true, '{1}'", idSubGrupo, this.txtBuscarCategoria.Text.ToString());
					this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);			
				}
			}
		}

		private void btMarcas_Click(object sender, System.EventArgs e)
		{
			using (ArtMarca miMarca = new ArtMarca())
			{
				if (DialogResult.OK == miMarca.ShowDialog())
				{					
					int idSubGrupo = (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value;
						
					string sSQL = string.Format("Exec ListaMarcas {0}, true, '{1}'", idSubGrupo, this.txtBuscarMarca.Text.ToString());
					this.ultraGrid3.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);			
				}
			}
		}

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtBuscaSubGrupo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{			
			int idGrupo = 0;
			if (drGrupo["idGrupoArticulo"] != System.DBNull.Value) idGrupo = (int)drGrupo["idGrupoArticulo"];

			string sSQL = string.Format("Exec ListaSubGrupos {0}, true, '{1}'", idGrupo, this.txtBuscaSubGrupo.Text.ToString());
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);	
		}

		private void txtBuscaSubGrupo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtBuscarCategoria_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtBuscarCategoria_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un subgrupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (this.ultraGrid1.ActiveRow.Cells["SubGrupo"].Value.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione un subgrupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string sSQL = string.Format("Exec ListaCategorias {0}, true, '{1}'", (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value, this.txtBuscarCategoria.Text.ToString());
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void txtBuscarMarca_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un subgrupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (this.ultraGrid1.ActiveRow.Cells["SubGrupo"].Value.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione un subgrupo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			string sSQL = string.Format("Exec ListaMarcas {0}, true, '{1}'", (int)this.ultraGrid1.ActiveRow.Cells["idSubGrupo"].Value, this.txtBuscarMarca.Text.ToString());
			this.ultraGrid3.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void txtBuscarMarca_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid3_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ArtGrupo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) 
			{
				if (this.barraDato1.bEditar)
					this.barraDato1_Deshacer(sender, e);
				else
					this.DialogResult = DialogResult.No;
			}
			
			if (e.KeyCode == Keys.N)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Crear(sender, e);
				}
			}
			if (e.KeyCode == Keys.E)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Editar(sender, e);
				}
			}
			if (e.KeyCode == Keys.B)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Buscar(sender, e);
				}
			}
			if (e.KeyCode == Keys.Delete)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Borrar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F12)
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Grabar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F5)
			{
				if (!this.barraDato1.bEditar)
				{
					barraDato1.IdRegistro = (int)drGrupo["idGrupoArticulo"];
					barraDato1.ProximoId(5);					
				}
			}
		}	
	}

}