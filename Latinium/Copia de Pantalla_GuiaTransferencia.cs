using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Pantalla_GuiaTransferencia.
	/// </summary>
	public class Pantalla_GuiaTransferencia : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPlaca;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor RucTransp;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTransportista;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRUCDest;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtLlegada;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDestinatario;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPuntoPartida;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEmision;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFinTraslado;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtIniTraslado;
		private System.Windows.Forms.ComboBox cmbTipoComprobant;
		private System.Windows.Forms.Label label13;
		public Infragistics.Win.UltraWinGrid.UltraGrid grdSerie;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsCompraS;
		private C1.Data.C1DataSet cdsCompra;
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataSet cdsCompraTransf;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Pantalla_GuiaTransferencia()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPlaca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.RucTransp = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtTransportista = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtRUCDest = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtLlegada = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.txtDestinatario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtPuntoPartida = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.dtFechaEmision = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.dtFinTraslado = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.dtIniTraslado = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbTipoComprobant = new System.Windows.Forms.ComboBox();
			this.label13 = new System.Windows.Forms.Label();
			this.grdSerie = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsCompraS = new C1.Data.C1DataSet();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdsCompraTransf = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.RucTransp)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransportista)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRUCDest)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLlegada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDestinatario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPuntoPartida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEmision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFinTraslado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIniTraslado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraS)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTransf)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(120, 448);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(120, 21);
			this.ultraTextEditor1.TabIndex = 74;
			// 
			// label11
			// 
			this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(56, 448);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(44, 16);
			this.label11.TabIndex = 73;
			this.label11.Text = "Numero";
			// 
			// ultraTextEditor2
			// 
			this.ultraTextEditor2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTextEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor2.Enabled = false;
			this.ultraTextEditor2.Location = new System.Drawing.Point(120, 424);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(120, 21);
			this.ultraTextEditor2.TabIndex = 72;
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(32, 424);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(73, 16);
			this.label12.TabIndex = 71;
			this.label12.Text = "Comprobante";
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Location = new System.Drawing.Point(368, 46);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.Size = new System.Drawing.Size(100, 21);
			this.txtIdCompra.TabIndex = 70;
			this.txtIdCompra.Text = "txtIdCompra";
			// 
			// txtPlaca
			// 
			this.txtPlaca.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtPlaca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPlaca.Location = new System.Drawing.Point(552, 112);
			this.txtPlaca.Name = "txtPlaca";
			this.txtPlaca.Size = new System.Drawing.Size(120, 21);
			this.txtPlaca.TabIndex = 69;
			// 
			// label10
			// 
			this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(512, 114);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(33, 16);
			this.label10.TabIndex = 68;
			this.label10.Text = "Placa";
			// 
			// RucTransp
			// 
			this.RucTransp.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.RucTransp.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.RucTransp.Location = new System.Drawing.Point(120, 400);
			this.RucTransp.Name = "RucTransp";
			this.RucTransp.Size = new System.Drawing.Size(120, 21);
			this.RucTransp.TabIndex = 67;
			// 
			// label9
			// 
			this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(48, 400);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(56, 16);
			this.label9.TabIndex = 66;
			this.label9.Text = "R.U.C./C.I";
			// 
			// txtTransportista
			// 
			this.txtTransportista.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTransportista.Location = new System.Drawing.Point(112, 112);
			this.txtTransportista.Name = "txtTransportista";
			this.txtTransportista.Size = new System.Drawing.Size(336, 21);
			this.txtTransportista.TabIndex = 65;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(16, 114);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(70, 16);
			this.label8.TabIndex = 64;
			this.label8.Text = "Transportista";
			// 
			// txtRUCDest
			// 
			this.txtRUCDest.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtRUCDest.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRUCDest.Location = new System.Drawing.Point(120, 376);
			this.txtRUCDest.Name = "txtRUCDest";
			this.txtRUCDest.Size = new System.Drawing.Size(336, 21);
			this.txtRUCDest.TabIndex = 63;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(72, 376);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(29, 16);
			this.label7.TabIndex = 62;
			this.label7.Text = "RUC";
			// 
			// txtLlegada
			// 
			this.txtLlegada.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtLlegada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLlegada.Location = new System.Drawing.Point(552, 80);
			this.txtLlegada.Name = "txtLlegada";
			this.txtLlegada.Size = new System.Drawing.Size(336, 21);
			this.txtLlegada.TabIndex = 61;
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(456, 82);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(89, 16);
			this.label6.TabIndex = 60;
			this.label6.Text = "Punto de llegada";
			// 
			// txtDestinatario
			// 
			this.txtDestinatario.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtDestinatario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDestinatario.Location = new System.Drawing.Point(120, 352);
			this.txtDestinatario.Name = "txtDestinatario";
			this.txtDestinatario.Size = new System.Drawing.Size(336, 21);
			this.txtDestinatario.TabIndex = 59;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(40, 352);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(68, 16);
			this.label5.TabIndex = 58;
			this.label5.Text = "Destinatario:";
			// 
			// txtPuntoPartida
			// 
			this.txtPuntoPartida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPuntoPartida.Location = new System.Drawing.Point(112, 80);
			this.txtPuntoPartida.Name = "txtPuntoPartida";
			this.txtPuntoPartida.Size = new System.Drawing.Size(336, 21);
			this.txtPuntoPartida.TabIndex = 57;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 82);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(87, 16);
			this.label4.TabIndex = 56;
			this.label4.Text = "Punto de partida";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 16);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(82, 16);
			this.label3.TabIndex = 46;
			this.label3.Text = "Fecha Emisión:";
			// 
			// dtFechaEmision
			// 
			this.dtFechaEmision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaEmision.DateButtons.Add(dateButton1);
			this.dtFechaEmision.Format = "dd/MMM/yyyy";
			this.dtFechaEmision.Location = new System.Drawing.Point(112, 14);
			this.dtFechaEmision.Name = "dtFechaEmision";
			this.dtFechaEmision.NonAutoSizeHeight = 23;
			this.dtFechaEmision.Size = new System.Drawing.Size(104, 21);
			this.dtFechaEmision.TabIndex = 45;
			this.dtFechaEmision.Value = new System.DateTime(2009, 10, 28, 0, 0, 0, 0);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(440, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(70, 16);
			this.label2.TabIndex = 44;
			this.label2.Text = "Fin Traslado:";
			// 
			// dtFinTraslado
			// 
			this.dtFinTraslado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFinTraslado.DateButtons.Add(dateButton2);
			this.dtFinTraslado.Format = "dd/MMM/yyyy";
			this.dtFinTraslado.Location = new System.Drawing.Point(512, 14);
			this.dtFinTraslado.Name = "dtFinTraslado";
			this.dtFinTraslado.NonAutoSizeHeight = 23;
			this.dtFinTraslado.Size = new System.Drawing.Size(104, 21);
			this.dtFinTraslado.TabIndex = 43;
			this.dtFinTraslado.Value = new System.DateTime(2009, 10, 28, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(232, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(81, 16);
			this.label1.TabIndex = 42;
			this.label1.Text = "Inicio Traslado:";
			// 
			// dtIniTraslado
			// 
			this.dtIniTraslado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtIniTraslado.DateButtons.Add(dateButton3);
			this.dtIniTraslado.Format = "dd/MMM/yyyy";
			this.dtIniTraslado.Location = new System.Drawing.Point(320, 14);
			this.dtIniTraslado.Name = "dtIniTraslado";
			this.dtIniTraslado.NonAutoSizeHeight = 23;
			this.dtIniTraslado.Size = new System.Drawing.Size(104, 21);
			this.dtIniTraslado.TabIndex = 41;
			this.dtIniTraslado.Value = new System.DateTime(2009, 10, 28, 0, 0, 0, 0);
			// 
			// cmbTipoComprobant
			// 
			this.cmbTipoComprobant.Items.AddRange(new object[] {
																													 "FACTURA",
																													 "CONSIGNACION",
																													 "TRASLADO ENTRE ESTABLECIMIENTOS DE UNA MISMA EMPRESA",
																													 "DEVOLUCION"});
			this.cmbTipoComprobant.Location = new System.Drawing.Point(144, 46);
			this.cmbTipoComprobant.Name = "cmbTipoComprobant";
			this.cmbTipoComprobant.Size = new System.Drawing.Size(216, 21);
			this.cmbTipoComprobant.TabIndex = 75;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(24, 48);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(121, 16);
			this.label13.TabIndex = 76;
			this.label13.Text = "Tipo COMPROBANTE:";
			// 
			// grdSerie
			// 
			this.grdSerie.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdSerie.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdSerie.DisplayLayout.Appearance = appearance1;
			this.grdSerie.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 105;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 123;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 139;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 167;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn5.Header.Caption = "# Documento";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 81;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 70;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 288;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 428;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.grdSerie.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdSerie.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grdSerie.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.grdSerie.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdSerie.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerie.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerie.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerie.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.grdSerie.DisplayLayout.UseFixedHeaders = true;
			this.grdSerie.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdSerie.Location = new System.Drawing.Point(16, 152);
			this.grdSerie.Name = "grdSerie";
			this.grdSerie.Size = new System.Drawing.Size(888, 184);
			this.grdSerie.TabIndex = 77;
			// 
			// ultraDataSource1
			// 
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// cdsCompraS
			// 
			this.cdsCompraS.BindingContextCtrl = this;
			this.cdsCompraS.DataLibrary = "LibFacturacion";
			this.cdsCompraS.DataLibraryUrl = "";
			this.cdsCompraS.DataSetDef = "dsCompraTabla";
			this.cdsCompraS.EnforceConstraints = false;
			this.cdsCompraS.FillOnRequest = false;
			this.cdsCompraS.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraS.Name = "cdsCompraS";
			this.cdsCompraS.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraS.SchemaDef = null;
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.EnforceConstraints = false;
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
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
			// 
			// cdsCompraTransf
			// 
			this.cdsCompraTransf.BindingContextCtrl = this;
			this.cdsCompraTransf.DataLibrary = "LibFacturacion";
			this.cdsCompraTransf.DataLibraryUrl = "";
			this.cdsCompraTransf.DataSetDef = "dsCompraTransf";
			this.cdsCompraTransf.EnforceConstraints = false;
			this.cdsCompraTransf.FillOnRequest = false;
			this.cdsCompraTransf.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompraTransf.Name = "cdsCompraTransf";
			this.cdsCompraTransf.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompraTransf.SchemaDef = null;
			// 
			// Pantalla_GuiaTransferencia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(920, 494);
			this.Controls.Add(this.grdSerie);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbTipoComprobant);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.ultraTextEditor2);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtPlaca);
			this.Controls.Add(this.RucTransp);
			this.Controls.Add(this.txtTransportista);
			this.Controls.Add(this.txtRUCDest);
			this.Controls.Add(this.txtLlegada);
			this.Controls.Add(this.txtDestinatario);
			this.Controls.Add(this.txtPuntoPartida);
			this.Controls.Add(this.dtFechaEmision);
			this.Controls.Add(this.dtFinTraslado);
			this.Controls.Add(this.dtIniTraslado);
			this.Name = "Pantalla_GuiaTransferencia";
			this.Text = "Impresión de GUIAS de REMISION y TRANSFERENCIAS";
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.RucTransp)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransportista)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRUCDest)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLlegada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDestinatario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPuntoPartida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEmision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFinTraslado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIniTraslado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraS)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompraTransf)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

	}
}
