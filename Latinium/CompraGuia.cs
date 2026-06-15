using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraGuia.
	/// </summary>
	public class CompraGuia : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsGuiaRem;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVenta;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor5;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor6;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor7;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor8;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor9;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPuntoPartida;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDestinatario;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtLlegada;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRUCDest;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTransportista;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor RucTransp;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPlaca;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdCompra;
		private C1.Data.C1DataView cdvGuiRem;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtIniTraslado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFinTraslado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEmision;
		private C1.Data.C1DataSet cdsSeteo;
		private C1.Data.C1DataSet cdsCliente;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor2;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.Misc.UltraButton btReportes;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraGuia()
		{
			InitializeComponent();
		}

		int IdCompra = -1;
		public CompraGuia(int idCompra)
		{
			InitializeComponent();
			IdCompra = idCompra;
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(CompraGuia));
			this.cdsGuiaRem = new C1.Data.C1DataSet();
			this.chkVenta = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtIniTraslado = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.dtFinTraslado = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.dtFechaEmision = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor4 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor5 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor6 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor7 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor8 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor9 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtPuntoPartida = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDestinatario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtLlegada = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.txtRUCDest = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtTransportista = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.RucTransp = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtPlaca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdvGuiRem = new C1.Data.C1DataView();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.ultraTextEditor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.btReportes = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRem)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIniTraslado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFinTraslado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEmision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPuntoPartida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDestinatario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLlegada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRUCDest)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransportista)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.RucTransp)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiRem)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsGuiaRem
			// 
			this.cdsGuiaRem.BindingContextCtrl = this;
			this.cdsGuiaRem.DataLibrary = "LibFacturacion";
			this.cdsGuiaRem.DataLibraryUrl = "";
			this.cdsGuiaRem.DataSetDef = "dsGuiaRemision";
			this.cdsGuiaRem.FillOnRequest = false;
			this.cdsGuiaRem.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsGuiaRem.Name = "cdsGuiaRem";
			this.cdsGuiaRem.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGuiaRem.SchemaDef = null;
			this.cdsGuiaRem.PositionChanged += new C1.Data.PositionChangeEventHandler(this.cdsGuiaRem_PositionChanged);
			this.cdsGuiaRem.BeforeFill += new C1.Data.FillEventHandler(this.cdsGuiaRem_BeforeFill);
			// 
			// chkVenta
			// 
			this.chkVenta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.chkVenta.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsGuiaRem, "GuiaRemision.Venta"));
			this.chkVenta.Location = new System.Drawing.Point(32, 64);
			this.chkVenta.Name = "chkVenta";
			this.chkVenta.Size = new System.Drawing.Size(72, 20);
			this.chkVenta.TabIndex = 1;
			this.chkVenta.Text = "Venta";
			// 
			// dtIniTraslado
			// 
			this.dtIniTraslado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtIniTraslado.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.FechaInicio"));
			dateButton1.Caption = "Today";
			this.dtIniTraslado.DateButtons.Add(dateButton1);
			this.dtIniTraslado.Format = "dd/MMM/yyyy";
			this.dtIniTraslado.Location = new System.Drawing.Point(128, 8);
			this.dtIniTraslado.Name = "dtIniTraslado";
			this.dtIniTraslado.NonAutoSizeHeight = 23;
			this.dtIniTraslado.Size = new System.Drawing.Size(104, 21);
			this.dtIniTraslado.TabIndex = 2;
			this.dtIniTraslado.Value = new System.DateTime(2009, 10, 28, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(32, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(81, 16);
			this.label1.TabIndex = 3;
			this.label1.Text = "Inicio Traslado:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(264, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(70, 16);
			this.label2.TabIndex = 5;
			this.label2.Text = "Fin Traslado:";
			// 
			// dtFinTraslado
			// 
			this.dtFinTraslado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFinTraslado.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.FechaFin"));
			dateButton2.Caption = "Today";
			this.dtFinTraslado.DateButtons.Add(dateButton2);
			this.dtFinTraslado.Format = "dd/MMM/yyyy";
			this.dtFinTraslado.Location = new System.Drawing.Point(360, 8);
			this.dtFinTraslado.Name = "dtFinTraslado";
			this.dtFinTraslado.NonAutoSizeHeight = 23;
			this.dtFinTraslado.Size = new System.Drawing.Size(104, 21);
			this.dtFinTraslado.TabIndex = 4;
			this.dtFinTraslado.Value = new System.DateTime(2009, 10, 28, 0, 0, 0, 0);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(32, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(82, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "Fecha Emisión:";
			// 
			// dtFechaEmision
			// 
			this.dtFechaEmision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaEmision.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.FechaEmision"));
			dateButton3.Caption = "Today";
			this.dtFechaEmision.DateButtons.Add(dateButton3);
			this.dtFechaEmision.Format = "dd/MMM/yyyy";
			this.dtFechaEmision.Location = new System.Drawing.Point(128, 40);
			this.dtFechaEmision.Name = "dtFechaEmision";
			this.dtFechaEmision.NonAutoSizeHeight = 23;
			this.dtFechaEmision.Size = new System.Drawing.Size(104, 21);
			this.dtFechaEmision.TabIndex = 6;
			this.dtFechaEmision.Value = new System.DateTime(2009, 10, 28, 0, 0, 0, 0);
			// 
			// ultraCheckEditor1
			// 
			this.ultraCheckEditor1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsGuiaRem, "GuiaRemision.Compra"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(32, 88);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(72, 20);
			this.ultraCheckEditor1.TabIndex = 8;
			this.ultraCheckEditor1.Text = "Compra";
			// 
			// ultraCheckEditor2
			// 
			this.ultraCheckEditor2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsGuiaRem, "GuiaRemision.Transformacion"));
			this.ultraCheckEditor2.Location = new System.Drawing.Point(32, 112);
			this.ultraCheckEditor2.Name = "ultraCheckEditor2";
			this.ultraCheckEditor2.Size = new System.Drawing.Size(104, 20);
			this.ultraCheckEditor2.TabIndex = 9;
			this.ultraCheckEditor2.Text = "Transformación";
			// 
			// ultraCheckEditor3
			// 
			this.ultraCheckEditor3.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor3.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsGuiaRem, "GuiaRemision.Consignacion"));
			this.ultraCheckEditor3.Location = new System.Drawing.Point(32, 136);
			this.ultraCheckEditor3.Name = "ultraCheckEditor3";
			this.ultraCheckEditor3.Size = new System.Drawing.Size(112, 20);
			this.ultraCheckEditor3.TabIndex = 10;
			this.ultraCheckEditor3.Text = "Consignación";
			// 
			// ultraCheckEditor4
			// 
			this.ultraCheckEditor4.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor4.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsGuiaRem, "GuiaRemision.Traslado2"));
			this.ultraCheckEditor4.Location = new System.Drawing.Point(168, 112);
			this.ultraCheckEditor4.Name = "ultraCheckEditor4";
			this.ultraCheckEditor4.Size = new System.Drawing.Size(184, 40);
			this.ultraCheckEditor4.TabIndex = 11;
			this.ultraCheckEditor4.Text = "Traslado por emisor itinerante de comprobantes de venta";
			// 
			// ultraCheckEditor5
			// 
			this.ultraCheckEditor5.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor5.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsGuiaRem, "GuiaRemision.Traslado1"));
			this.ultraCheckEditor5.Location = new System.Drawing.Point(168, 72);
			this.ultraCheckEditor5.Name = "ultraCheckEditor5";
			this.ultraCheckEditor5.Size = new System.Drawing.Size(192, 32);
			this.ultraCheckEditor5.TabIndex = 12;
			this.ultraCheckEditor5.Text = "Traslado entre establecimientos de una misma empresa";
			// 
			// ultraCheckEditor6
			// 
			this.ultraCheckEditor6.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor6.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsGuiaRem, "GuiaRemision.Devolucion"));
			this.ultraCheckEditor6.Location = new System.Drawing.Point(376, 64);
			this.ultraCheckEditor6.Name = "ultraCheckEditor6";
			this.ultraCheckEditor6.Size = new System.Drawing.Size(88, 20);
			this.ultraCheckEditor6.TabIndex = 13;
			this.ultraCheckEditor6.Text = "Devolución";
			// 
			// ultraCheckEditor7
			// 
			this.ultraCheckEditor7.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor7.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsGuiaRem, "GuiaRemision.Importacion"));
			this.ultraCheckEditor7.Location = new System.Drawing.Point(376, 88);
			this.ultraCheckEditor7.Name = "ultraCheckEditor7";
			this.ultraCheckEditor7.Size = new System.Drawing.Size(88, 20);
			this.ultraCheckEditor7.TabIndex = 14;
			this.ultraCheckEditor7.Text = "Importación";
			// 
			// ultraCheckEditor8
			// 
			this.ultraCheckEditor8.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor8.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsGuiaRem, "GuiaRemision.Exportacion"));
			this.ultraCheckEditor8.Location = new System.Drawing.Point(376, 112);
			this.ultraCheckEditor8.Name = "ultraCheckEditor8";
			this.ultraCheckEditor8.Size = new System.Drawing.Size(96, 20);
			this.ultraCheckEditor8.TabIndex = 15;
			this.ultraCheckEditor8.Text = "Exportación";
			// 
			// ultraCheckEditor9
			// 
			this.ultraCheckEditor9.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraCheckEditor9.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsGuiaRem, "GuiaRemision.Otros"));
			this.ultraCheckEditor9.Location = new System.Drawing.Point(376, 136);
			this.ultraCheckEditor9.Name = "ultraCheckEditor9";
			this.ultraCheckEditor9.Size = new System.Drawing.Size(72, 20);
			this.ultraCheckEditor9.TabIndex = 16;
			this.ultraCheckEditor9.Text = "Otros";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(32, 168);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(87, 16);
			this.label4.TabIndex = 17;
			this.label4.Text = "Punto de partida";
			// 
			// txtPuntoPartida
			// 
			this.txtPuntoPartida.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.Partida"));
			this.txtPuntoPartida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPuntoPartida.Location = new System.Drawing.Point(128, 168);
			this.txtPuntoPartida.Name = "txtPuntoPartida";
			this.txtPuntoPartida.Size = new System.Drawing.Size(336, 21);
			this.txtPuntoPartida.TabIndex = 18;
			// 
			// txtDestinatario
			// 
			this.txtDestinatario.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.NombreDestinatario"));
			this.txtDestinatario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDestinatario.Location = new System.Drawing.Point(128, 192);
			this.txtDestinatario.Name = "txtDestinatario";
			this.txtDestinatario.Size = new System.Drawing.Size(336, 21);
			this.txtDestinatario.TabIndex = 20;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(32, 192);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(68, 16);
			this.label5.TabIndex = 19;
			this.label5.Text = "Destinatario:";
			// 
			// txtLlegada
			// 
			this.txtLlegada.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.Llegada"));
			this.txtLlegada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLlegada.Location = new System.Drawing.Point(128, 216);
			this.txtLlegada.Name = "txtLlegada";
			this.txtLlegada.Size = new System.Drawing.Size(336, 21);
			this.txtLlegada.TabIndex = 22;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(32, 216);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(89, 16);
			this.label6.TabIndex = 21;
			this.label6.Text = "Punto de llegada";
			// 
			// txtRUCDest
			// 
			this.txtRUCDest.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.RucDestinatario"));
			this.txtRUCDest.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRUCDest.Location = new System.Drawing.Point(128, 240);
			this.txtRUCDest.Name = "txtRUCDest";
			this.txtRUCDest.Size = new System.Drawing.Size(336, 21);
			this.txtRUCDest.TabIndex = 24;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(32, 240);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(29, 16);
			this.label7.TabIndex = 23;
			this.label7.Text = "RUC";
			// 
			// txtTransportista
			// 
			this.txtTransportista.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.TranspNombre"));
			this.txtTransportista.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTransportista.Location = new System.Drawing.Point(128, 264);
			this.txtTransportista.Name = "txtTransportista";
			this.txtTransportista.Size = new System.Drawing.Size(336, 21);
			this.txtTransportista.TabIndex = 26;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(32, 264);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(70, 16);
			this.label8.TabIndex = 25;
			this.label8.Text = "Transportista";
			// 
			// RucTransp
			// 
			this.RucTransp.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.TranspRuc"));
			this.RucTransp.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.RucTransp.Location = new System.Drawing.Point(128, 288);
			this.RucTransp.Name = "RucTransp";
			this.RucTransp.Size = new System.Drawing.Size(120, 21);
			this.RucTransp.TabIndex = 28;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(32, 288);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(70, 16);
			this.label9.TabIndex = 27;
			this.label9.Text = "RUC Transp.";
			// 
			// txtPlaca
			// 
			this.txtPlaca.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.Placa"));
			this.txtPlaca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPlaca.Location = new System.Drawing.Point(344, 288);
			this.txtPlaca.Name = "txtPlaca";
			this.txtPlaca.Size = new System.Drawing.Size(120, 21);
			this.txtPlaca.TabIndex = 30;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(280, 288);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(33, 16);
			this.label10.TabIndex = 29;
			this.label10.Text = "Placa";
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(120, 352);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 31;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.Location = new System.Drawing.Point(320, 352);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 32;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.idCompra"));
			this.txtIdCompra.Location = new System.Drawing.Point(368, 32);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.Size = new System.Drawing.Size(100, 21);
			this.txtIdCompra.TabIndex = 33;
			this.txtIdCompra.Text = "txtIdCompra";
			// 
			// cdvGuiRem
			// 
			this.cdvGuiRem.BindingContextCtrl = this;
			this.cdvGuiRem.DataSet = this.cdsGuiaRem;
			this.cdvGuiRem.Sort = "idCompra";
			this.cdvGuiRem.TableName = "";
			this.cdvGuiRem.TableViewName = "GuiaRemision";
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.BarraMovimiento = true;
			this.barraDatoSQL1.DataMember = "GuiaRemision";
			this.barraDatoSQL1.DataNombreId = "idGuiaRemision";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsGuiaRem;
			this.barraDatoSQL1.DataTabla = "GuiaRemision";
			this.barraDatoSQL1.DataTablaHija = "";
			this.barraDatoSQL1.Location = new System.Drawing.Point(24, 352);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(304, 24);
			this.barraDatoSQL1.TabIndex = 34;
			this.barraDatoSQL1.Visible = false;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = false;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Load += new System.EventHandler(this.barraDatoSQL1_Load);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.Numero"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(344, 312);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(120, 21);
			this.ultraTextEditor1.TabIndex = 39;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(280, 312);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(44, 16);
			this.label11.TabIndex = 38;
			this.label11.Text = "Numero";
			// 
			// ultraTextEditor2
			// 
			this.ultraTextEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsGuiaRem, "GuiaRemision.Factura"));
			this.ultraTextEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor2.Enabled = false;
			this.ultraTextEditor2.Location = new System.Drawing.Point(128, 312);
			this.ultraTextEditor2.Name = "ultraTextEditor2";
			this.ultraTextEditor2.Size = new System.Drawing.Size(120, 21);
			this.ultraTextEditor2.TabIndex = 37;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(32, 312);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(73, 16);
			this.label12.TabIndex = 36;
			this.label12.Text = "Comprobante";
			// 
			// btReportes
			// 
			this.btReportes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.Image = ((object)(resources.GetObject("appearance1.Image")));
			this.btReportes.Appearance = appearance1;
			this.btReportes.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btReportes.Location = new System.Drawing.Point(408, 352);
			this.btReportes.Name = "btReportes";
			this.btReportes.Size = new System.Drawing.Size(24, 23);
			this.btReportes.TabIndex = 40;
			this.btReportes.Click += new System.EventHandler(this.btReportes_Click);
			// 
			// CompraGuia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(496, 390);
			this.Controls.Add(this.btReportes);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.ultraTextEditor2);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.txtPlaca);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.RucTransp);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.txtTransportista);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtRUCDest);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtLlegada);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtDestinatario);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtPuntoPartida);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.ultraCheckEditor9);
			this.Controls.Add(this.ultraCheckEditor8);
			this.Controls.Add(this.ultraCheckEditor7);
			this.Controls.Add(this.ultraCheckEditor6);
			this.Controls.Add(this.ultraCheckEditor5);
			this.Controls.Add(this.ultraCheckEditor4);
			this.Controls.Add(this.ultraCheckEditor3);
			this.Controls.Add(this.ultraCheckEditor2);
			this.Controls.Add(this.ultraCheckEditor1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtFechaEmision);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.dtFinTraslado);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtIniTraslado);
			this.Controls.Add(this.chkVenta);
			this.Controls.Add(this.barraDatoSQL1);
			this.MaximizeBox = false;
			this.Name = "CompraGuia";
			this.Text = "Guia de Remisión";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.CompraGuia_Closing);
			this.Load += new System.EventHandler(this.CompraGuia_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsGuiaRem)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIniTraslado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFinTraslado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEmision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPuntoPartida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDestinatario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLlegada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRUCDest)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTransportista)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.RucTransp)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGuiRem)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			string stComprobante = Funcion.NumeraFactura(cdsCliente, 30);
			if (stComprobante.Length > 0)
			{
				drGuia["Numero"] = stComprobante;
			}
			cdsGuiaRem.Update();
			Close();
		}

		C1.Data.C1DataRow drGuia;
		public DateTime dtFecha = DateTime.Today;
		public int IdCliente = 0;
		private void CompraGuia_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Guia Remision'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
			Cursor = Cursors.Default;
			this.txtIdCompra.Width = 0;
			if (IdCompra > 0) 
			{
				string stSelect = "Select idGuiaRemision From GuiaRemision Where idCompra = " + IdCompra.ToString();
				int IdGuia = Funcion.iEscalarSQL(cdsGuiaRem, stSelect, false);
				if (IdGuia > 0)
				{
					barraDatoSQL1.HabilitaControles(false);
					barraDatoSQL1.Visible = true;
					barraDatoSQL1.IdRegistro = IdGuia;
					barraDatoSQL1.ProximoId(5);
				
					btAceptar.Visible  = false;
					btCancelar.Visible  = false;
				}
				else
				{
					BindingContext[this.cdsGuiaRem,"GuiaRemision"].AddNew();
					C1.Data.C1DataRow drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];
					C1.Data.C1DataRow drCliente = cdsCliente.TableViews["Cliente"].Rows.Find(IdCliente);
					drGuia["idCompra"] = IdCompra;
					drGuia["Partida"] = drSeteo["Direccion"];
					drGuia["FechaInicio"] = dtFecha;
					drGuia["FechaFin"] = dtFecha;
					drGuia["FechaEmision"] = dtFecha;
//					drGuia["LLegada"] = drCliente["Direccion"];
//					drGuia["NombreDestinatario"] = drCliente["Nombre"];
//					drGuia["RucDestinatario"] = drCliente["RUC"];
					string stNumero = "Select Numero From Compra "
						+ "Where idCompra = " + IdCompra.ToString();
					stNumero = Funcion.sEscalarSQL(cdsCliente, stNumero, false);
					drGuia["Factura"] = stNumero;
				}
			}
		}

		private void cdsGuiaRem_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsGuiaRem.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsGuiaRem_PositionChanged(object sender, C1.Data.PositionChangeEventArgs e)
		{
			drGuia = e.Row;
		}

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Eliminar la Guía de Remisión?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			barraDatoSQL1.BorraRegistro();
		}

		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			MessageBox.Show("Guias se crean desde las facturas", "Información");
		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			barraDatoSQL1.DeshacerRegistro();
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			barraDatoSQL1.EditarRegistro();
		}

		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			barraDatoSQL1.GrabaRegistro();
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Compra.idCompra} = " + txtIdCompra.Value.ToString();
			using (Reporte miReporte = new Reporte("GuiaRemision.Rpt", stFiltro))
			{
				miReporte.ShowDialog();
			}
			Cursor = Cursors.Default;
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void CompraGuia_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Guia Remision'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}

		private void btReportes_Click(object sender, System.EventArgs e)
		{
			CompraGuiaImp miGuia = new CompraGuiaImp();
			miGuia.MdiParent = this.MdiParent;
			miGuia.Show();
		}

		private void barraDatoSQL1_Load(object sender, System.EventArgs e)
		{
		
		}

	
	}
}
