using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de MarceloImportacion.
	/// </summary>
	public class MarceloImportacion : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFob;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFobExw;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFleteExterior;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSeguro;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAdvalorem;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCif;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFodinfa;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAlmacenaje;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCertificadoInen;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTasaSeguro;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDesconsolidacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFleteInterno;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHonorarios;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private C1.Data.C1DataSet cdsImportacion;
		private C1.Data.C1DataView cdvImportacion;
		private System.Windows.Forms.Label label34;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSalvaguardia;
		private System.Windows.Forms.Label label35;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAdvSalvaguardia;
		private System.Windows.Forms.Label label26;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtOAE;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txt;
		private System.Windows.Forms.Label label27;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtParesZ;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label ff;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtArancel;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSalva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLib;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtOtros;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstibaje;
		private Infragistics.Win.Misc.UltraButton btImprimir;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label lbltotal;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MarceloImportacion()
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
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			this.label38 = new System.Windows.Forms.Label();
			this.txtFob = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtFobExw = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFleteExterior = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSeguro = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFodinfa = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtAdvalorem = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtCif = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtAlmacenaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCertificadoInen = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTasaSeguro = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.txtDesconsolidacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtFleteInterno = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtHonorarios = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.btImprimir = new Infragistics.Win.Misc.UltraButton();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsImportacion = new C1.Data.C1DataSet();
			this.cdvImportacion = new C1.Data.C1DataView();
			this.txtSalvaguardia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label34 = new System.Windows.Forms.Label();
			this.label35 = new System.Windows.Forms.Label();
			this.txtAdvSalvaguardia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.txtOAE = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txt = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label27 = new System.Windows.Forms.Label();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtParesZ = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.ff = new System.Windows.Forms.Label();
			this.txtArancel = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSalva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtLib = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.txtOtros = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.txtEstibaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.lbltotal = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtFob)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFobExw)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFleteExterior)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSeguro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFodinfa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdvalorem)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCif)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAlmacenaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCertificadoInen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTasaSeguro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesconsolidacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFleteInterno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHonorarios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalvaguardia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdvSalvaguardia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOAE)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtParesZ)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArancel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLib)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstibaje)).BeginInit();
			this.SuspendLayout();
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.Location = new System.Drawing.Point(109, 48);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(27, 16);
			this.label38.TabIndex = 10;
			this.label38.Text = "FOB";
			// 
			// txtFob
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFob.Appearance = appearance1;
			this.txtFob.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Fob"));
			this.txtFob.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFob.FormatString = "#,##0.00";
			this.txtFob.Location = new System.Drawing.Point(144, 48);
			this.txtFob.Name = "txtFob";
			this.txtFob.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFob.PromptChar = ' ';
			this.txtFob.Size = new System.Drawing.Size(104, 21);
			this.txtFob.TabIndex = 11;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(49, 16);
			this.label1.TabIndex = 13;
			this.label1.Text = "Proyecto";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(80, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 16);
			this.label2.TabIndex = 14;
			this.label2.Text = "FOB EXW";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(11, 96);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(125, 16);
			this.label3.TabIndex = 15;
			this.label3.Text = "FLETE DEL EXTERIOR";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(83, 120);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(53, 16);
			this.label4.TabIndex = 16;
			this.label4.Text = "SEGURO";
			// 
			// txtFobExw
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFobExw.Appearance = appearance2;
			this.txtFobExw.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.FobExw"));
			this.txtFobExw.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFobExw.FormatString = "#,##0.00";
			this.txtFobExw.Location = new System.Drawing.Point(144, 72);
			this.txtFobExw.Name = "txtFobExw";
			this.txtFobExw.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFobExw.PromptChar = ' ';
			this.txtFobExw.Size = new System.Drawing.Size(104, 21);
			this.txtFobExw.TabIndex = 17;
			// 
			// txtFleteExterior
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFleteExterior.Appearance = appearance3;
			this.txtFleteExterior.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.FleteExt"));
			this.txtFleteExterior.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFleteExterior.FormatString = "#,##0.00";
			this.txtFleteExterior.Location = new System.Drawing.Point(144, 96);
			this.txtFleteExterior.Name = "txtFleteExterior";
			this.txtFleteExterior.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFleteExterior.PromptChar = ' ';
			this.txtFleteExterior.Size = new System.Drawing.Size(104, 21);
			this.txtFleteExterior.TabIndex = 18;
			// 
			// txtSeguro
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSeguro.Appearance = appearance4;
			this.txtSeguro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Seguro"));
			this.txtSeguro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSeguro.FormatString = "#,##0.00";
			this.txtSeguro.Location = new System.Drawing.Point(144, 120);
			this.txtSeguro.Name = "txtSeguro";
			this.txtSeguro.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSeguro.PromptChar = ' ';
			this.txtSeguro.Size = new System.Drawing.Size(104, 21);
			this.txtSeguro.TabIndex = 19;
			// 
			// txtIva
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance5;
			this.txtIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.IvaImpor"));
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(144, 264);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(104, 21);
			this.txtIva.TabIndex = 27;
			// 
			// txtFodinfa
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFodinfa.Appearance = appearance6;
			this.txtFodinfa.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Fodinfa"));
			this.txtFodinfa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFodinfa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFodinfa.FormatString = "#,##0.00";
			this.txtFodinfa.Location = new System.Drawing.Point(144, 288);
			this.txtFodinfa.Name = "txtFodinfa";
			this.txtFodinfa.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFodinfa.PromptChar = ' ';
			this.txtFodinfa.Size = new System.Drawing.Size(104, 21);
			this.txtFodinfa.TabIndex = 26;
			// 
			// txtAdvalorem
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAdvalorem.Appearance = appearance7;
			this.txtAdvalorem.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Advalorem1"));
			this.txtAdvalorem.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAdvalorem.FormatString = "#,##0.00";
			this.txtAdvalorem.Location = new System.Drawing.Point(144, 192);
			this.txtAdvalorem.Name = "txtAdvalorem";
			this.txtAdvalorem.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAdvalorem.PromptChar = ' ';
			this.txtAdvalorem.Size = new System.Drawing.Size(104, 21);
			this.txtAdvalorem.TabIndex = 25;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(27, 272);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(109, 16);
			this.label5.TabIndex = 24;
			this.label5.Text = "IVA IMPORTACION";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(81, 296);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(55, 16);
			this.label6.TabIndex = 23;
			this.label6.Text = "FODINFA";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(51, 192);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(85, 16);
			this.label7.TabIndex = 22;
			this.label7.Text = "ADVALOREM 1";
			// 
			// txtCif
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCif.Appearance = appearance8;
			this.txtCif.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Cif"));
			this.txtCif.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCif.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCif.FormatString = "#,##0.00";
			this.txtCif.Location = new System.Drawing.Point(144, 168);
			this.txtCif.Name = "txtCif";
			this.txtCif.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCif.PromptChar = ' ';
			this.txtCif.Size = new System.Drawing.Size(104, 21);
			this.txtCif.TabIndex = 21;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(113, 168);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(26, 16);
			this.label8.TabIndex = 20;
			this.label8.Text = "CIF ";
			// 
			// txtAlmacenaje
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAlmacenaje.Appearance = appearance9;
			this.txtAlmacenaje.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Almacenaje"));
			this.txtAlmacenaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAlmacenaje.FormatString = "#,##0.00";
			this.txtAlmacenaje.Location = new System.Drawing.Point(144, 400);
			this.txtAlmacenaje.Name = "txtAlmacenaje";
			this.txtAlmacenaje.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAlmacenaje.PromptChar = ' ';
			this.txtAlmacenaje.Size = new System.Drawing.Size(104, 21);
			this.txtAlmacenaje.TabIndex = 35;
			// 
			// txtCertificadoInen
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCertificadoInen.Appearance = appearance10;
			this.txtCertificadoInen.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Inen"));
			this.txtCertificadoInen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCertificadoInen.FormatString = "#,##0.00";
			this.txtCertificadoInen.Location = new System.Drawing.Point(144, 352);
			this.txtCertificadoInen.Name = "txtCertificadoInen";
			this.txtCertificadoInen.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCertificadoInen.PromptChar = ' ';
			this.txtCertificadoInen.Size = new System.Drawing.Size(104, 21);
			this.txtCertificadoInen.TabIndex = 34;
			// 
			// txtTasaSeguro
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTasaSeguro.Appearance = appearance11;
			this.txtTasaSeguro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.TasaSeguro"));
			this.txtTasaSeguro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTasaSeguro.FormatString = "#,##0.00";
			this.txtTasaSeguro.Location = new System.Drawing.Point(144, 144);
			this.txtTasaSeguro.Name = "txtTasaSeguro";
			this.txtTasaSeguro.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTasaSeguro.PromptChar = ' ';
			this.txtTasaSeguro.Size = new System.Drawing.Size(104, 21);
			this.txtTasaSeguro.TabIndex = 33;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(56, 408);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 16);
			this.label9.TabIndex = 32;
			this.label9.Text = "ALMACENAJE";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(24, 352);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(111, 16);
			this.label10.TabIndex = 31;
			this.label10.Text = "CERTIFICADO INEN";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(24, 144);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(112, 16);
			this.label11.TabIndex = 30;
			this.label11.Text = "TASAS DE SEGURO";
			// 
			// txtDesconsolidacion
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDesconsolidacion.Appearance = appearance12;
			this.txtDesconsolidacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Desconsolidacion"));
			this.txtDesconsolidacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDesconsolidacion.FormatString = "#,##0.00";
			this.txtDesconsolidacion.Location = new System.Drawing.Point(144, 328);
			this.txtDesconsolidacion.Name = "txtDesconsolidacion";
			this.txtDesconsolidacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDesconsolidacion.PromptChar = ' ';
			this.txtDesconsolidacion.Size = new System.Drawing.Size(104, 21);
			this.txtDesconsolidacion.TabIndex = 29;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(8, 328);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(123, 16);
			this.label12.TabIndex = 28;
			this.label12.Text = "DESCONSOLIDACION";
			// 
			// txtFleteInterno
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFleteInterno.Appearance = appearance13;
			this.txtFleteInterno.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.FleteInterno"));
			this.txtFleteInterno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFleteInterno.FormatString = "#,##0.00";
			this.txtFleteInterno.Location = new System.Drawing.Point(144, 496);
			this.txtFleteInterno.Name = "txtFleteInterno";
			this.txtFleteInterno.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFleteInterno.PromptChar = ' ';
			this.txtFleteInterno.Size = new System.Drawing.Size(104, 21);
			this.txtFleteInterno.TabIndex = 41;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(40, 504);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(93, 16);
			this.label15.TabIndex = 38;
			this.label15.Text = "FLETE INTERNO";
			// 
			// txtHonorarios
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtHonorarios.Appearance = appearance14;
			this.txtHonorarios.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Honorarios"));
			this.txtHonorarios.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHonorarios.FormatString = "#,##0.00";
			this.txtHonorarios.Location = new System.Drawing.Point(144, 472);
			this.txtHonorarios.Name = "txtHonorarios";
			this.txtHonorarios.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtHonorarios.PromptChar = ' ';
			this.txtHonorarios.Size = new System.Drawing.Size(104, 21);
			this.txtHonorarios.TabIndex = 37;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(0, 480);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(135, 16);
			this.label16.TabIndex = 36;
			this.label16.Text = "HON. TRAMITE ADUANA";
			// 
			// btImprimir
			// 
			this.btImprimir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btImprimir.Location = new System.Drawing.Point(712, 24);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(88, 23);
			this.btImprimir.TabIndex = 80;
			this.btImprimir.Text = "Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btTraer_Click);
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible =1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			this.cdsProyecto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// cmbBodega
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance15;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.idImportacion"));
			this.cmbBodega.DataMember = "SubProyecto";
			this.cmbBodega.DataSource = this.cdsProyecto;
			this.cmbBodega.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Proyecto";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 49;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 10;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 230;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 15;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 54;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 15;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 55;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 17;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(80, 16);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(248, 21);
			this.cmbBodega.TabIndex = 98;
			this.cmbBodega.ValueMember = "idSubProyecto";
			this.cmbBodega.Validated += new System.EventHandler(this.cmbBodega_Validated);
			this.cmbBodega.RowSelected += new Infragistics.Win.UltraWinGrid.RowSelectedEventHandler(this.cmbBodega_RowSelected);
			// 
			// cdsImportacion
			// 
			this.cdsImportacion.BindingContextCtrl = this;
			this.cdsImportacion.DataLibrary = "LibFacturacion";
			this.cdsImportacion.DataLibraryUrl = "";
			this.cdsImportacion.DataSetDef = "dsMarceloImportacion";
			this.cdsImportacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsImportacion.Name = "cdsImportacion";
			this.cdsImportacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsImportacion.SchemaDef = null;
			this.cdsImportacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsImportacion_BeforeFill);
			// 
			// cdvImportacion
			// 
			this.cdvImportacion.BindingContextCtrl = this;
			this.cdvImportacion.DataSet = this.cdsImportacion;
			this.cdvImportacion.TableName = "MarceloImportacion";
			this.cdvImportacion.TableViewName = "";
			// 
			// txtSalvaguardia
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSalvaguardia.Appearance = appearance16;
			this.txtSalvaguardia.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Advalorem2"));
			this.txtSalvaguardia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSalvaguardia.FormatString = "#,##0.00";
			this.txtSalvaguardia.Location = new System.Drawing.Point(144, 216);
			this.txtSalvaguardia.Name = "txtSalvaguardia";
			this.txtSalvaguardia.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSalvaguardia.PromptChar = ' ';
			this.txtSalvaguardia.Size = new System.Drawing.Size(104, 21);
			this.txtSalvaguardia.TabIndex = 99;
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.Location = new System.Drawing.Point(44, 224);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(92, 16);
			this.label34.TabIndex = 101;
			this.label34.Text = "SALVAGUARDIA";
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label35.Location = new System.Drawing.Point(55, 248);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(81, 16);
			this.label35.TabIndex = 102;
			this.label35.Text = "ADVALOREM ";
			// 
			// txtAdvSalvaguardia
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAdvSalvaguardia.Appearance = appearance17;
			this.txtAdvSalvaguardia.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Advalorem"));
			this.txtAdvSalvaguardia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAdvSalvaguardia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtAdvSalvaguardia.FormatString = "#,##0.00";
			this.txtAdvSalvaguardia.Location = new System.Drawing.Point(144, 240);
			this.txtAdvSalvaguardia.Name = "txtAdvSalvaguardia";
			this.txtAdvSalvaguardia.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAdvSalvaguardia.PromptChar = ' ';
			this.txtAdvSalvaguardia.Size = new System.Drawing.Size(104, 21);
			this.txtAdvSalvaguardia.TabIndex = 103;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.Location = new System.Drawing.Point(104, 376);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(28, 16);
			this.label26.TabIndex = 104;
			this.label26.Text = "OAE";
			// 
			// txtOAE
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOAE.Appearance = appearance18;
			this.txtOAE.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.OAE"));
			this.txtOAE.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOAE.FormatString = "#,##0.00";
			this.txtOAE.Location = new System.Drawing.Point(144, 376);
			this.txtOAE.Name = "txtOAE";
			this.txtOAE.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtOAE.PromptChar = ' ';
			this.txtOAE.Size = new System.Drawing.Size(104, 21);
			this.txtOAE.TabIndex = 105;
			// 
			// txt
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txt.Appearance = appearance19;
			this.txt.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Demoraje"));
			this.txt.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txt.FormatString = "#,##0.00";
			this.txt.Location = new System.Drawing.Point(144, 424);
			this.txt.Name = "txt";
			this.txt.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txt.PromptChar = ' ';
			this.txt.Size = new System.Drawing.Size(104, 21);
			this.txt.TabIndex = 106;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.Location = new System.Drawing.Point(64, 432);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(67, 16);
			this.label27.TabIndex = 107;
			this.label27.Text = "DEMORAJE";
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "MarceloImportacion";
			this.barraDato1.DataNombreId = "idImportacion";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsImportacion;
			this.barraDato1.DataTabla = "MarceloImportacion";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.barraDato1.Location = new System.Drawing.Point(480, 504);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 108;
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
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtParesZ
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtParesZ.Appearance = appearance20;
			this.txtParesZ.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.ParesZ"));
			this.txtParesZ.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtParesZ.FormatString = "";
			this.txtParesZ.Location = new System.Drawing.Point(616, 24);
			this.txtParesZ.Name = "txtParesZ";
			this.txtParesZ.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtParesZ.PromptChar = ' ';
			this.txtParesZ.Size = new System.Drawing.Size(64, 21);
			this.txtParesZ.TabIndex = 109;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(512, 32);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(98, 16);
			this.label13.TabIndex = 110;
			this.label13.Text = "PARES ZAPATOS";
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(512, 56);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(43, 16);
			this.label14.TabIndex = 111;
			this.label14.Text = "Arancel";
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(464, 80);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(151, 16);
			this.label17.TabIndex = 112;
			this.label17.Text = "Salvaguardia x par de zapato";
			// 
			// ff
			// 
			this.ff.AutoSize = true;
			this.ff.Location = new System.Drawing.Point(464, 104);
			this.ff.Name = "ff";
			this.ff.Size = new System.Drawing.Size(136, 16);
			this.ff.TabIndex = 113;
			this.ff.Text = "Por. descuento Liberación";
			// 
			// txtArancel
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArancel.Appearance = appearance21;
			this.txtArancel.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Arancel"));
			this.txtArancel.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArancel.FormatString = "";
			this.txtArancel.Location = new System.Drawing.Point(616, 48);
			this.txtArancel.Name = "txtArancel";
			this.txtArancel.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtArancel.PromptChar = ' ';
			this.txtArancel.Size = new System.Drawing.Size(64, 21);
			this.txtArancel.TabIndex = 114;
			// 
			// txtSalva
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSalva.Appearance = appearance22;
			this.txtSalva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.PorSalvaguardia"));
			this.txtSalva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSalva.FormatString = "";
			this.txtSalva.Location = new System.Drawing.Point(616, 72);
			this.txtSalva.Name = "txtSalva";
			this.txtSalva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSalva.PromptChar = ' ';
			this.txtSalva.Size = new System.Drawing.Size(64, 21);
			this.txtSalva.TabIndex = 115;
			this.txtSalva.ValueChanged += new System.EventHandler(this.txtSalva_ValueChanged);
			// 
			// txtLib
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLib.Appearance = appearance23;
			this.txtLib.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Liberacion"));
			this.txtLib.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLib.FormatString = "";
			this.txtLib.Location = new System.Drawing.Point(616, 96);
			this.txtLib.Name = "txtLib";
			this.txtLib.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtLib.PromptChar = ' ';
			this.txtLib.Size = new System.Drawing.Size(64, 21);
			this.txtLib.TabIndex = 116;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(88, 528);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(44, 16);
			this.label18.TabIndex = 117;
			this.label18.Text = "OTROS";
			// 
			// txtOtros
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOtros.Appearance = appearance24;
			this.txtOtros.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Otros"));
			this.txtOtros.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOtros.FormatString = "#,##0.00";
			this.txtOtros.Location = new System.Drawing.Point(144, 520);
			this.txtOtros.Name = "txtOtros";
			this.txtOtros.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtOtros.PromptChar = ' ';
			this.txtOtros.Size = new System.Drawing.Size(104, 21);
			this.txtOtros.TabIndex = 118;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(64, 456);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(58, 16);
			this.label19.TabIndex = 119;
			this.label19.Text = "ESTIBAJE";
			// 
			// txtEstibaje
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstibaje.Appearance = appearance25;
			this.txtEstibaje.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "MarceloImportacion.Estibaje"));
			this.txtEstibaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstibaje.FormatString = "#,##0.00";
			this.txtEstibaje.Location = new System.Drawing.Point(144, 448);
			this.txtEstibaje.Name = "txtEstibaje";
			this.txtEstibaje.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtEstibaje.PromptChar = ' ';
			this.txtEstibaje.Size = new System.Drawing.Size(104, 21);
			this.txtEstibaje.TabIndex = 120;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Location = new System.Drawing.Point(456, 152);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(98, 16);
			this.label20.TabIndex = 121;
			this.label20.Text = "Total Importacion :";
			this.label20.Click += new System.EventHandler(this.label20_Click);
			// 
			// lbltotal
			// 
			this.lbltotal.AutoSize = true;
			this.lbltotal.Location = new System.Drawing.Point(568, 152);
			this.lbltotal.Name = "lbltotal";
			this.lbltotal.Size = new System.Drawing.Size(98, 16);
			this.lbltotal.TabIndex = 122;
			this.lbltotal.Text = "Total Importacion :";
			// 
			// MarceloImportacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(904, 549);
			this.Controls.Add(this.lbltotal);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.txtEstibaje);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.txtOtros);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.txtLib);
			this.Controls.Add(this.txtSalva);
			this.Controls.Add(this.txtArancel);
			this.Controls.Add(this.ff);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.txtParesZ);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.label27);
			this.Controls.Add(this.txt);
			this.Controls.Add(this.txtOAE);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.txtAdvSalvaguardia);
			this.Controls.Add(this.label35);
			this.Controls.Add(this.label34);
			this.Controls.Add(this.txtSalvaguardia);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.txtFleteInterno);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.txtHonorarios);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.txtAlmacenaje);
			this.Controls.Add(this.txtCertificadoInen);
			this.Controls.Add(this.txtTasaSeguro);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.txtDesconsolidacion);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtFodinfa);
			this.Controls.Add(this.txtAdvalorem);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtCif);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtSeguro);
			this.Controls.Add(this.txtFleteExterior);
			this.Controls.Add(this.txtFobExw);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtFob);
			this.Controls.Add(this.label38);
			this.Name = "MarceloImportacion";
			this.Text = "Marcelo Importacion";
			this.Load += new System.EventHandler(this.MarceloImportacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtFob)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFobExw)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFleteExterior)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSeguro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFodinfa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdvalorem)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCif)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAlmacenaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCertificadoInen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTasaSeguro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesconsolidacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFleteInterno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHonorarios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalvaguardia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdvSalvaguardia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOAE)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtParesZ)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArancel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLib)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOtros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstibaje)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void MarceloImportacion_Load(object sender, System.EventArgs e)
		{
			
			this.barraDato1.PosUltima();
			this.barraDato1.HabilitaControles(false);
			this.cmbBodega.Enabled=true;  
	
				
  
		}

  	private void btTraer_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			//string stReporte = "MarceloImportacion.rpt";
			string stFiltro = "{Compra.idSubproyecto} = " + this.cmbBodega.Value.ToString();
			stFiltro += " And {DetCompra.idImportacion}=1  ";
			Reporte miReporte;
			miReporte = new Reporte("MarceloImportacion", stFiltro );
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();

		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsProyecto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{

		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();  
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			barraDato1.EditarRegistro();  
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
		barraDato1.GrabaRegistro();  
		}

		private void cmbBodega_Validated(object sender, System.EventArgs e)
		{
			Funcion.EjecutaSQL(cdsImportacion , "exec MarceloImportacionP " + (int)cmbBodega.Value);  

			barraDato1.IdRegistro =  (int)cmbBodega.Value;
			barraDato1.ProximoId(5);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
		
		}

		private void cdsImportacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsImportacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cmbBodega_RowSelected(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
		{
			Funcion.EjecutaSQL(cdsImportacion , "exec MarceloImportacionP " + (int)cmbBodega.Value);  
			barraDato1.IdRegistro =  (int)cmbBodega.Value;
			barraDato1.ProximoId(5);
		}

		private void txtSalva_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label20_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
