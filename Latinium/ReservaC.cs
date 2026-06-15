using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ReservaC.
	/// </summary>
	public class ReservaC : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsReserva;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArt1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArt2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArt3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArt4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumero;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ReservaC()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			this.cdsReserva = new C1.Data.C1DataSet();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cmbArt1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cmbArt2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbArt3 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbArt4 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsReserva
			// 
			this.cdsReserva.BindingContextCtrl = this;
			this.cdsReserva.DataLibrary = "LibFacturacion";
			this.cdsReserva.DataLibraryUrl = "";
			this.cdsReserva.DataSetDef = "dsReservacion";
			this.cdsReserva.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReserva.Name = "cdsReserva";
			this.cdsReserva.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsReserva.SchemaDef = null;
			this.cdsReserva.BeforeFill += new C1.Data.FillEventHandler(this.cdsReserva_BeforeFill);
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cmbArt1
			// 
			this.cmbArt1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArt1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArt1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReserva, "ReservaC.idArticulo1"));
			this.cmbArt1.DataSource = this.cdvArticulo;
			this.cmbArt1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 38;
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 34;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Width = 86;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 64;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 47;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbArt1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbArt1.DisplayMember = "Articulo";
			this.cmbArt1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArt1.Location = new System.Drawing.Point(120, 24);
			this.cmbArt1.Name = "cmbArt1";
			this.cmbArt1.Size = new System.Drawing.Size(216, 21);
			this.cmbArt1.TabIndex = 0;
			this.cmbArt1.ValueMember = "idArticulo";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(45, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Servicio";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(24, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(79, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Intermediación";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 120);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(62, 16);
			this.label3.TabIndex = 5;
			this.label3.Text = "Adicionales";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(24, 88);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(56, 16);
			this.label4.TabIndex = 7;
			this.label4.Text = "Impuestos";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(56, 192);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 11;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(232, 192);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 12;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cmbArt2
			// 
			this.cmbArt2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArt2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArt2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReserva, "ReservaC.idArticulo2"));
			this.cmbArt2.DataSource = this.cdvArticulo;
			this.cmbArt2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 38;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 34;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Width = 86;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Width = 64;
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Width = 47;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbArt2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbArt2.DisplayMember = "Articulo";
			this.cmbArt2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArt2.Location = new System.Drawing.Point(120, 56);
			this.cmbArt2.Name = "cmbArt2";
			this.cmbArt2.Size = new System.Drawing.Size(216, 21);
			this.cmbArt2.TabIndex = 13;
			this.cmbArt2.ValueMember = "idArticulo";
			// 
			// cmbArt3
			// 
			this.cmbArt3.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArt3.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArt3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReserva, "ReservaC.idArticulo3"));
			this.cmbArt3.DataSource = this.cdvArticulo;
			this.cmbArt3.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 38;
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 34;
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Width = 86;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 7;
			ultraGridColumn26.Width = 64;
			ultraGridColumn27.Header.VisiblePosition = 8;
			ultraGridColumn27.Width = 47;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbArt3.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbArt3.DisplayMember = "Articulo";
			this.cmbArt3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArt3.Location = new System.Drawing.Point(120, 88);
			this.cmbArt3.Name = "cmbArt3";
			this.cmbArt3.Size = new System.Drawing.Size(216, 21);
			this.cmbArt3.TabIndex = 14;
			this.cmbArt3.ValueMember = "idArticulo";
			// 
			// cmbArt4
			// 
			this.cmbArt4.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArt4.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArt4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReserva, "ReservaC.idArticulo4"));
			this.cmbArt4.DataSource = this.cdvArticulo;
			this.cmbArt4.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 3;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 5;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 38;
			ultraGridColumn31.Header.VisiblePosition = 6;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 34;
			ultraGridColumn32.Header.VisiblePosition = 4;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Width = 86;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 7;
			ultraGridColumn35.Width = 64;
			ultraGridColumn36.Header.VisiblePosition = 8;
			ultraGridColumn36.Width = 47;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36});
			this.cmbArt4.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbArt4.DisplayMember = "Articulo";
			this.cmbArt4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArt4.Location = new System.Drawing.Point(120, 120);
			this.cmbArt4.Name = "cmbArt4";
			this.cmbArt4.Size = new System.Drawing.Size(216, 21);
			this.cmbArt4.TabIndex = 15;
			this.cmbArt4.ValueMember = "idArticulo";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(24, 152);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(44, 16);
			this.label5.TabIndex = 16;
			this.label5.Text = "Numero";
			// 
			// txtNumero
			// 
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsReserva, "ReservaC.Numero"));
			this.txtNumero.Location = new System.Drawing.Point(120, 152);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.PromptChar = ' ';
			this.txtNumero.Size = new System.Drawing.Size(100, 21);
			this.txtNumero.TabIndex = 17;
			// 
			// ReservaC
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(368, 238);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbArt4);
			this.Controls.Add(this.cmbArt3);
			this.Controls.Add(this.cmbArt2);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbArt1);
			this.Name = "ReservaC";
			this.Text = "Reservación Constantes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ReservaC_Closing);
			this.Load += new System.EventHandler(this.ReservaC_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsReserva.Update();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Grabar Constantes de Reservacion");
			}
			Close();
		}

		private void cdsReserva_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsReserva.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ReservaC_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'ReservaConst'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void ReservaC_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'ReservaConst'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}
	}
}
