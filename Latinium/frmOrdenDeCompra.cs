using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using Latinium.Models.OrdenDeCompra.Bodega;
using Latinium.Services.OrdenDeCompra.Bodega;
using Latinium.Services.OrdenDeCompra.Excel;
using Latinium.Services.OrdenDeCompra.Importacion;
using Latinium.Models.OrdenDeCompra.Importacion;
using System.IO;
using System.Text;
using System.Net;
using Latinium.Services.OrdenDeCompra.Configuracion;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmOrdenDeCompra.
	/// </summary>
	public class frmOrdenDeCompra : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdOrdenCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSubTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGrupoCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdComprobante;
		private System.ComponentModel.IContainer components;
		private DetalleBodegaService _bodegaService = new DetalleBodegaService();
		private System.Windows.Forms.Button btnPlantilla;
		private System.Windows.Forms.Button btnImportar;
		private System.Windows.Forms.ToolTip toolTip1;

		int IdOC = 0;
		private bool procesandoImportacion = false;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConsignacion;

		private enum EstadoFormulario
		{
			Inicial = 0,
			Nuevo = 1,
			Consulta = 2,
			Edicion = 3,
			Procesando = 4
		}

		private EstadoFormulario estadoFormulario =
			EstadoFormulario.Inicial;

		public frmOrdenDeCompra()
		{
			InitializeComponent();
			this.KeyPreview = true;
		}

		public frmOrdenDeCompra(int idOC) : this()
		{
			IdOC = idOC;
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmOrdenDeCompra));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UIO");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GYE");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnBodega", 0);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GYE");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			this.label2 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSubTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdOrdenCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoRuc = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdGrupoCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.txtIdComprobante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnImportar = new System.Windows.Forms.Button();
			this.btnPlantilla = new System.Windows.Forms.Button();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.chkConsignacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdOrdenCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(224, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 373;
			this.label2.Text = "Número";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 18);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 372;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 16);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 374;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// txtNumero
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance2;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(280, 16);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 21);
			this.txtNumero.TabIndex = 371;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(16, 56);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(29, 16);
			this.lblIdentificacion.TabIndex = 378;
			this.lblIdentificacion.Text = "RUC";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(224, 56);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 16);
			this.lblCliente.TabIndex = 377;
			this.lblCliente.Text = "Nombre";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance3;
			this.txtRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(72, 54);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 21);
			this.txtRuc.TabIndex = 376;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			// 
			// txtNombre
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance4;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(280, 56);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(240, 21);
			this.txtNombre.TabIndex = 375;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 280);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 16);
			this.label3.TabIndex = 381;
			this.label3.Text = "Notas";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance5;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 304);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(424, 96);
			this.txtNotas.TabIndex = 380;
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(-128, 408);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1032, 8);
			this.gbBotones.TabIndex = 382;
			this.gbBotones.TabStop = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 424);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 383;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnNuevo, "Crear una nueva orden de compra");
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(88, 424);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 387;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnConsultar, "Consultar órdenes de compra registradas");
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(174, 424);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 386;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnEditar, "Editar la orden de compra seleccionada");
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(258, 424);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 384;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnGuardar, "Guardar cambios de la orden de compra");
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(594, 424);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 388;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnCancelar, "Cancelar la operación actual");
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ForeColor = System.Drawing.Color.Black;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(678, 424);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 389;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnSalir, "Cerrar la ventana");
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(342, 424);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 390;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnAnular, "Anular la orden de compra actual");
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(552, 376);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 16);
			this.label1.TabIndex = 402;
			this.label1.Text = "TOTAL";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(552, 320);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(64, 16);
			this.lblSubTotal.TabIndex = 395;
			this.lblSubTotal.Text = "SUBTOTAL";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(552, 348);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 394;
			this.lblIva.Text = "IVA";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance6;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(632, 368);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(128, 28);
			this.txtTotal.TabIndex = 401;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance7;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(680, 340);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 21);
			this.txtIvaTotal.TabIndex = 400;
			// 
			// txtSubTotal
			// 
			this.txtSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSubTotal.Appearance = appearance8;
			this.txtSubTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSubTotal.Enabled = false;
			this.txtSubTotal.FormatString = "#,##0.00";
			this.txtSubTotal.Location = new System.Drawing.Point(680, 312);
			this.txtSubTotal.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtSubTotal.Name = "txtSubTotal";
			this.txtSubTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSubTotal.PromptChar = ' ';
			this.txtSubTotal.Size = new System.Drawing.Size(80, 21);
			this.txtSubTotal.TabIndex = 397;
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
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(744, 272);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(16, 21);
			this.txtIdCliente.TabIndex = 403;
			this.txtIdCliente.Visible = false;
			// 
			// txtIdOrdenCompra
			// 
			this.txtIdOrdenCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdOrdenCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdOrdenCompra.Enabled = false;
			this.txtIdOrdenCompra.Location = new System.Drawing.Point(728, 272);
			this.txtIdOrdenCompra.Name = "txtIdOrdenCompra";
			this.txtIdOrdenCompra.PromptChar = ' ';
			this.txtIdOrdenCompra.Size = new System.Drawing.Size(16, 21);
			this.txtIdOrdenCompra.TabIndex = 404;
			this.txtIdOrdenCompra.Visible = false;
			// 
			// txtIdTipoRuc
			// 
			this.txtIdTipoRuc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoRuc.Enabled = false;
			this.txtIdTipoRuc.Location = new System.Drawing.Point(712, 272);
			this.txtIdTipoRuc.Name = "txtIdTipoRuc";
			this.txtIdTipoRuc.PromptChar = ' ';
			this.txtIdTipoRuc.Size = new System.Drawing.Size(16, 21);
			this.txtIdTipoRuc.TabIndex = 405;
			this.txtIdTipoRuc.Visible = false;
			// 
			// txtIdGrupoCliente
			// 
			this.txtIdGrupoCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGrupoCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGrupoCliente.Enabled = false;
			this.txtIdGrupoCliente.Location = new System.Drawing.Point(696, 272);
			this.txtIdGrupoCliente.Name = "txtIdGrupoCliente";
			this.txtIdGrupoCliente.PromptChar = ' ';
			this.txtIdGrupoCliente.Size = new System.Drawing.Size(16, 21);
			this.txtIdGrupoCliente.TabIndex = 406;
			this.txtIdGrupoCliente.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance9.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance9;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "_DetCompra";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Artículo";
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 152;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 125;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 260;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance10;
			ultraGridColumn5.Format = "#,##0";
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.MaskInput = "";
			ultraGridColumn5.NullText = "0";
			ultraGridColumn5.PromptChar = ' ';
			ultraGridColumn5.Width = 55;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance11;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn6.PromptChar = ' ';
			ultraGridColumn6.Width = 104;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance12;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.PromptChar = ' ';
			ultraGridColumn7.Width = 113;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance13;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 37;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance14;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 37;
			appearance15.Image = ((object)(resources.GetObject("appearance15.Image")));
			appearance15.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance15;
			appearance16.Image = ((object)(resources.GetObject("appearance16.Image")));
			appearance16.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellButtonAppearance = appearance16;
			ultraGridColumn10.Header.Caption = "Asignación";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn10.Width = 82;
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
																										 ultraGridColumn10});
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance17;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance18;
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 8.5F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance20;
			appearance21.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance21.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance23;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 96);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(760, 168);
			this.ultraGrid1.TabIndex = 379;
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn5.DefaultValue = 0;
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn6.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(576, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 407;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(680, 272);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 408;
			this.txtEstado.Visible = false;
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BackColor = System.Drawing.Color.Transparent;
			this.lblUsuario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(576, 48);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 26);
			this.lblUsuario.TabIndex = 409;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtIdComprobante
			// 
			this.txtIdComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdComprobante.Enabled = false;
			this.txtIdComprobante.Location = new System.Drawing.Point(664, 272);
			this.txtIdComprobante.Name = "txtIdComprobante";
			this.txtIdComprobante.PromptChar = ' ';
			this.txtIdComprobante.Size = new System.Drawing.Size(16, 21);
			this.txtIdComprobante.TabIndex = 410;
			this.txtIdComprobante.Visible = false;
			// 
			// btnImportar
			// 
			this.btnImportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImportar.CausesValidation = false;
			this.btnImportar.Image = ((System.Drawing.Image)(resources.GetObject("btnImportar.Image")));
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(426, 424);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.TabIndex = 663;
			this.btnImportar.Text = "Importar";
			this.btnImportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnImportar, "Importar detalle de productos desde Excel");
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// btnPlantilla
			// 
			this.btnPlantilla.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnPlantilla.CausesValidation = false;
			this.btnPlantilla.Image = ((System.Drawing.Image)(resources.GetObject("btnPlantilla.Image")));
			this.btnPlantilla.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnPlantilla.Location = new System.Drawing.Point(510, 424);
			this.btnPlantilla.Name = "btnPlantilla";
			this.btnPlantilla.TabIndex = 664;
			this.btnPlantilla.Text = "Plantilla";
			this.btnPlantilla.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnPlantilla, "Descargar plantilla Excel para importación");
			this.btnPlantilla.Click += new System.EventHandler(this.btnPlantilla_Click);
			// 
			// toolTip1
			// 
			this.toolTip1.AutoPopDelay = 5000;
			this.toolTip1.InitialDelay = 300;
			this.toolTip1.ReshowDelay = 100;
			this.toolTip1.ShowAlways = true;
			// 
			// chkConsignacion
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkConsignacion.Appearance = appearance24;
			this.chkConsignacion.BackColor = System.Drawing.Color.Transparent;
			this.chkConsignacion.CausesValidation = false;
			this.chkConsignacion.Enabled = false;
			this.chkConsignacion.Location = new System.Drawing.Point(408, 16);
			this.chkConsignacion.Name = "chkConsignacion";
			this.chkConsignacion.Size = new System.Drawing.Size(96, 22);
			this.chkConsignacion.TabIndex = 665;
			this.chkConsignacion.Text = "Consignación";
			// 
			// frmOrdenDeCompra
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(768, 462);
			this.Controls.Add(this.chkConsignacion);
			this.Controls.Add(this.btnPlantilla);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.txtIdComprobante);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtIdGrupoCliente);
			this.Controls.Add(this.txtIdTipoRuc);
			this.Controls.Add(this.txtIdOrdenCompra);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtSubTotal);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.btnGuardar);
			this.Name = "frmOrdenDeCompra";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Orden De Compra Madre";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmOrdenDeCompra_KeyDown);
			this.Load += new System.EventHandler(this.frmOrdenDeCompra_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdOrdenCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		bool bEliminaAlValidar = false;
		private Acceso miAcceso;
		bool bRastreo = false;
		int IdSigno = 1;
		private C1DataRow drSeteoF;
		private bool bActualiza = true;
		decimal dPorcIva = 0.00m;

		private void frmOrdenDeCompra_Load(
			object sender,
			System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0636");

			if (!Funcion.Permiso("2179", cdsSeteoF))
			{
				MessageBox.Show(
					"No Puede Ingresar a Ordenes de Compra Madre",
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Stop);

				BeginInvoke(
					new MethodInvoker(this.Close));

				return;
			}

			this.dtFecha.CalendarInfo.MaxDate =
				DateTime.Today;

			drSeteoF =
				cdsSeteoF.TableViews["SeteoF"].Rows[0];

			bRastreo =
				Funcion.bEjecutaSQL(
				cdsSeteoF,
				"Exec SeteoGExiste 'RASTREO'"
				);

			IdSigno =
				Funcion.iEscalarSQL(
				cdsSeteoF,
				"SELECT Signo FROM CompraNumero WHERE idTipoFactura = 2"
				);

			dPorcIva =
				Funcion.decEscalarSQL(
				cdsSeteoF,
				string.Format(
				"Select dbo.RetornaImpuestoIVAVigente('{0}', 1)",
				DateTime.Today.ToString("yyyyMMdd")
				)
				);

			this.ultraGrid1.DisplayLayout.AutoFitColumns =
				true;

			FuncionesProcedimientos.EstadoGrids(
				false,
				this.ultraGrid1
				);

			// =========================
			// NUEVO
			// =========================

			if (IdOC > 0)
			{
				this.ConsultaMaestro(IdOC);

				CambiarEstadoFormulario(
					EstadoFormulario.Consulta);
			}
			else
			{
				CambiarEstadoFormulario(
					EstadoFormulario.Inicial);
			}
		}
		private void cmdTotal()
		{
			#region Creacion de Variables
			int dRedondeo = (int)drSeteoF["DecRedondeo"];

			dPorcIva = Funcion.decEscalarSQL(
				cdsSeteoF,
				string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)",
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

			decimal ivaRate = dPorcIva;
			if (ivaRate > 1.0m) ivaRate = ivaRate / 100.0m;
			#endregion Creacion de Variables

			this.txtIvaTotal.Value = 0.00m;
			this.txtSubTotal.Value = 0.00m;

			// Detecta qué columna de subtotal existe
			string subKey = "SubTotal";
			if (!ultraGrid1.DisplayLayout.Bands[0].Columns.Exists("SubTotal") &&
				ultraGrid1.DisplayLayout.Bands[0].Columns.Exists("Subtotal"))
			{
				subKey = "Subtotal";
			}

			#region Desplazamiento de grilla
			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["Cantidad"].Value != DBNull.Value &&	dr.Cells["Precio"].Value != DBNull.Value)
				{
					decimal cantidad = Convert.ToDecimal(dr.Cells["Cantidad"].Value);
					decimal precio= Convert.ToDecimal(dr.Cells["Precio"].Value);
					decimal sub= Math.Round(cantidad * precio, dRedondeo);

					dr.Cells[subKey].Value = sub;

					this.txtSubTotal.Value = Math.Round(Convert.ToDecimal(this.txtSubTotal.Value) + sub,dRedondeo);
				}
			}
			this.txtIvaTotal.Value = Math.Round(Convert.ToDecimal(this.txtSubTotal.Value) * ivaRate,dRedondeo	);

			this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtSubTotal.Value) + Convert.ToDecimal(this.txtIvaTotal.Value),dRedondeo);
			#endregion Desplazamiento de grilla
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private int ToInt(object value)
		{
			if (value == null || value == DBNull.Value)
				return 0;

			string s = value.ToString().Trim();
			if (s == "")
				return 0;

			return Convert.ToInt32(s);
		}

		private decimal ToDecimal(object value)
		{
			if (value == null || value == DBNull.Value)
				return 0;

			string s = value.ToString().Trim();
			if (s == "")
				return 0;

			return Convert.ToDecimal(s);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Elimina Filas Vacias
		
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid1, "idArticulo");
			bEliminaAlValidar = false;

			
			#endregion Elimina Filas Vacias

			if (!ValidaCantidadesUIOGYE()) return;
			this.cmdTotal();		
			#region Valida Estado 
			if (!bNuevo)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
				{
					int idDetCompra = 0;
					if (dr.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)dr.Cells["idDetCompra"].Value;
					int idArticulo = 0;
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)dr.Cells["idArticulo"].Value;						
				}		
			}
			#endregion Valida Estado 
			
			if (!Validacion.ValidaTotales((decimal)this.txtTotal.Value)) return;

			if (!Validacion.vNotas(this.txtNotas)) return;
			
			if (!Validacion.vbClienteExiste(this.txtRuc, true, 2, 6, cdsSeteoF)) return;
							
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "Compra", cdsSeteoF)) return;
				
			if (!Validacion.vbGrids(this.ultraGrid1, "Codigo", 2)) return;	

			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				decimal precio = (dr.Cells["Precio"].Value == System.DBNull.Value) ? 0m : Convert.ToDecimal(dr.Cells["Precio"].Value);
				if (precio <= 0m)
				{
					string cod = (dr.Cells.Exists("Codigo") && dr.Cells["Codigo"].Value != System.DBNull.Value) ? dr.Cells["Codigo"].Value.ToString().Trim() : "";
					string msg = "El precio no puede ser 0.";
					if (cod.Length > 0) msg = "El precio del artículo con código " + cod + " no puede ser 0.";
					MessageBox.Show(msg, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					try
					{
						this.ultraGrid1.ActiveRow = dr;
						if (dr.Cells.Exists("Precio")) this.ultraGrid1.ActiveCell = dr.Cells["Precio"];
					}
					catch { }

					return;
				}
			}


			if (DialogResult.Yes == MessageBox.Show("Antes de Guardar. \n\n VERIQUE QUE LOS DATOS REGISTRADOS SON LOS CORRECTOS \n\n INFORMACIÓN DEL PROVEEDOR \n PRODUCTOS \n PRECIOS CANTIDADES DESCUENTOS \n\n\n ¿Esta seguro de Guardar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 120;
						
					try
					{							
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;
						
						#region Numeracion
						if (bNuevo)
						{							
							oCmdActualiza.CommandText = "Exec FacturaNumeroEstacion 2";
							this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString(); 						 
						}
						#endregion Numeracion

						#region Maestro

						DateTime fechaTentativaLlegada = DateTime.Now.AddDays(60);
						DateTime horaLlegada = DateTime.Now;

						int consignacion = this.chkConsignacion.Checked ? 1 : 0;

						// si aún no manejas estos valores, los dejamos en 0
						int bFechaLlegada = 1;
						int bHoraLlegada = 1;
						int idCredTributario = 0;
						int idOrigen = 0;
						int idTipoFactura = 0;
						int idImportacionTipo = 0;
						int contadoCredito = 0;

						string sSQLMaestro = string.Format(
							"Exec MaestroOrdenDeCompra {0}, '{1}', '{2}', {3}, '{4}', {5}, {6}, {7}, {8}, " +
							"'{9}', '{10}', {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}",

							// =========================
							// CAMPOS ORIGINALES (9)
							// =========================
							(int)this.txtIdOrdenCompra.Value,
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),
							this.txtNumero.Text.ToString(),
							(int)this.txtIdCliente.Value,
							this.txtNotas.Text.ToString(),
							(decimal)this.txtSubTotal.Value,
							(decimal)this.txtIvaTotal.Value,
							(decimal)this.txtTotal.Value,
							(int)this.txtIdComprobante.Value,

							// =========================
							// NUEVOS CAMPOS (SP)
							// =========================
							fechaTentativaLlegada.ToString("yyyy-MM-dd HH:mm:ss"),
							horaLlegada.ToString("yyyy-MM-dd HH:mm:ss"),
							bFechaLlegada,
							bHoraLlegada,
							idCredTributario,
							idOrigen,
							idTipoFactura,
							idImportacionTipo,
							contadoCredito,
							consignacion
							);

						oCmdActualiza.CommandText = sSQLMaestro;
						this.txtIdOrdenCompra.Value = (int)oCmdActualiza.ExecuteScalar();

						#endregion Maestro

						#region Detalle

						DetalleBodegaService service =
							new DetalleBodegaService(); 

						foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{
							try
							{
								int idDetCompra = ToInt(dr.Cells["idDetCompra"].Value);
								int idOrden = ToInt(this.txtIdOrdenCompra.Value);
								int idArticulo = ToInt(dr.Cells["idArticulo"].Value);
								int cantidad = ToInt(dr.Cells["Cantidad"].Value);
								decimal precio = ToDecimal(dr.Cells["Precio"].Value);
								int uio = ToInt(dr.Cells["UIO"].Value);
								int gye = ToInt(dr.Cells["GYE"].Value);
								decimal subtotal = ToDecimal(dr.Cells["Subtotal"].Value);

								string sSQLDetalle = string.Format(
									"Exec GuardaDetOrdenCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}",
									idDetCompra,
									idOrden,
									idArticulo,
									cantidad,
									precio,
									uio,
									gye,
									subtotal
									);

								oCmdActualiza.CommandText = sSQLDetalle;

								object resDet = oCmdActualiza.ExecuteScalar();

								if (resDet == null || resDet == DBNull.Value)
									throw new Exception("No se pudo obtener el ID del detalle.");

								int idDetOrden = Convert.ToInt32(resDet);

								// BODEGAS
								service.GuardarDesdeTag(idDetOrden, dr, oCmdActualiza);

							}
							catch (Exception ex)
							{
								MessageBox.Show(
									"Error en detalle fila:\n" +
									"idArticulo: " + dr.Cells["idArticulo"].Value + "\n" +
									"Cantidad: " + dr.Cells["Cantidad"].Value + "\n" +
									"Precio: " + dr.Cells["Precio"].Value + "\n\n" +
									ex.Message,
									"Error conversión",
									MessageBoxButtons.OK,
									MessageBoxIcon.Error
									);
							}
						}
						#endregion Detalle			
				
						oTransaction.Commit();											

						int idCompraNuevo = (int)this.txtIdOrdenCompra.Value;
						this.btnCancelar_Click(sender, e);
						this.ConsultaMaestro(idCompraNuevo);

						CambiarEstadoFormulario(
							EstadoFormulario.Consulta);
					
						
					}
					catch (Exception ex)
					{
						try
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
						}
					}
					finally
					{
						oConexion.Close();
					}
				}
			}
		}

		private void btnConsultar_Click(
			object sender,
			System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaOrdenDeCompra Busqueda =
							 new frmBuscaOrdenDeCompra())
			{
				if (DialogResult.OK ==
					Busqueda.ShowDialog())
				{
					if (
						Busqueda.uGridOrdenesDeCompra.ActiveRow
						!= null
						)
					{
						this.ConsultaMaestro(
							(int)Busqueda
							.uGridOrdenesDeCompra
							.ActiveRow
							.Cells["IdOrdenDeCompra"]
							.Value);

						CambiarEstadoFormulario(
							EstadoFormulario.Consulta);
					}
				}
				else
				{
					CambiarEstadoFormulario(
						EstadoFormulario.Inicial);
				}
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.txtIdOrdenCompra.Value = 0;
			this.txtIdCliente.Value = 0;		

			this.txtRuc.Text = "";
			this.txtNombre.Text = "";		
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFecha.CalendarInfo.MinDate = Funcion.dtEscalarSQL(cdsSeteoF, "Select CONVERT(Date, FechaInicio) From SeteoF");			
			this.txtNumero.Text = "";
			this.chkConsignacion.Checked = false;
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleOrdenDeCompra 0");

			this.txtNotas.Text = "";		
			this.txtSubTotal.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;

			bAnulado = false;
			bNuevo = false;
			bEdicion = false;
			bEliminaAlValidar = false;			
			this.txtRuc.Enabled = false;
			this.txtNombre.Enabled = false;			
			this.dtFecha.Enabled = false;		
			this.txtNumero.Enabled = false;	
			
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			
			this.txtNotas.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;			
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;						
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Text = "Anular";
		
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			//this.ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.AllowEdit;

			CambiarEstadoFormulario(
				EstadoFormulario.Inicial);

		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(true, 2);
				if (sIdentificacion.Length > 0)
				{						
					this.txtRuc.Text = sIdentificacion;
					
					CargaCliente();
					
				}				
			}
			if (e.KeyCode == Keys.Enter)
			{
				this.ultraGrid1.Focus();
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}			
		}
		private void CargaCliente()
		{
			string[] sParametros = new string[11];
			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(true, 2, this.txtRuc.Text.ToString(), cdsSeteoF);

			if (sParametros[0] != null) this.txtIdCliente.Value = int.Parse(sParametros[0]);
			if (sParametros[1] != null) this.txtIdTipoRuc.Value = int.Parse(sParametros[1]);
			if (sParametros[2] != null) this.txtIdGrupoCliente.Value = sParametros[2];
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion) return;

			if (ultraGrid1.ActiveCell != null)
			{
				FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);

				if (e.KeyCode == Keys.F3)
				{
					try
					{
						string col = ultraGrid1.ActiveCell.Column.Key.ToUpper();

						if (col == "CODIGO" || col == "ARTICULO")
						{
							int idDetCompra = 0;
							int idArticulo = 0;
							DateTime fecha = DateTime.Today;

							if (ultraGrid1.ActiveRow.Cells["idDetCompra"].Value != null &&
								ultraGrid1.ActiveRow.Cells["idDetCompra"].Value != DBNull.Value)
								idDetCompra = Convert.ToInt32(ultraGrid1.ActiveRow.Cells["idDetCompra"].Value);

							if (ultraGrid1.ActiveRow.Cells["idArticulo"].Value != null &&
								ultraGrid1.ActiveRow.Cells["idArticulo"].Value != DBNull.Value)
								idArticulo = Convert.ToInt32(ultraGrid1.ActiveRow.Cells["idArticulo"].Value);

							if (this.dtFecha != null && this.dtFecha.Value != null && this.dtFecha.Value != DBNull.Value)
								fecha = Convert.ToDateTime(this.dtFecha.Value);

							FuncionesInventario.AbreBusquedaArticulos(this.ultraGrid1, 1, 2, 6, fecha, false);
						}
					}
					catch (Exception ex)
					{
						MessageBox.Show(ex.Message, "Error en Ingreso de Productos", 
							MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.txtEstado.Value = 1;						
			
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
		
			this.txtRuc.Enabled = true;
			this.txtNotas.Enabled = true;
		
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
	
			this.txtRuc.Focus();

			bNuevo = true;
			bEdicion = true;

			CambiarEstadoFormulario(
				EstadoFormulario.Nuevo);
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (!bEdicion || !bActualiza) return;
			bActualiza = false;

			try
			{
				#region Verifica Cambio			
				#region Codigo o Articulo
				if (e.Cell.Column.ToString() == "Codigo")
				{
					#region Variables
					string stCod = "";
					int IdArt = 0;
					int iCuenta = 0;
					decimal dPrecio = 0.00m;
					if (bRastreo) Rastreo("Inicio Cambio de " + e.Cell.Column.ToString());
					#endregion Variables
					#region Codigo
					if (e.Cell.Column.ToString() == "Codigo")
					{
						if (e.Cell.Row.Cells["Codigo"].Value == DBNull.Value || 
							e.Cell.Row.Cells["Codigo"].Value.ToString().Length == 0)
						{
							e.Cell.Row.Cells["Codigo"].Value = e.Cell.OriginalValue.ToString();
						}

						stCod = e.Cell.Row.Cells["Codigo"].Value.ToString();						
					}
					#endregion Codigo
					#region Articulo
					if (e.Cell.Column.ToString() == "Articulo")
					{
						if (e.Cell.Row.Cells["Articulo"].Value != DBNull.Value)
							stCod = e.Cell.Row.Cells["Articulo"].Value.ToString();
					}
					#endregion Articulo

					int iTipoArticulo = 0;
								
					#region Cuenta coincidencias del codigo
					string stSelect = string.Format("Exec VentaArticuloCuenta '{0}', {1}, {2}, {3}, 0", stCod, 2, 0, iTipoArticulo);
					if (bRastreo) Rastreo(String.Format("Cuenta Coincidencias de {0}: {1}", e.Cell.Column.ToString(), stCod));
					iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
					#endregion Cuenta coincidencias del codigo
					#region Mas de una coincidencia o ninguna muestra la lista
					if (iCuenta != 1)
					{ 
						if (bRastreo) Rastreo(string.Format("# de Coincidencias != de 1 en {0} : {1}", e.Cell.Column.ToString(), stCod));
					
						int Bodega = 0;

						using (BuscaArticulos miBusqueda = new BuscaArticulos(stCod, Bodega, 2))
						{
							if (DialogResult.OK == miBusqueda.ShowDialog())
							{
								if (bRastreo) Rastreo("Articulo escogido de la lista");
								IdArt = miBusqueda.idArticulo;
							}
							else
							{
								e.Cell.Row.Cells["idArticulo"].Value = 0;
								e.Cell.Row.Cells["Precio"].Value = 0.00m;
								e.Cell.Row.Cells["Codigo"].Value = "";
								e.Cell.Row.Cells["Articulo"].Value = "";
								cmdTotal();

								if (bRastreo) Rastreo("Cancela la selección de Artículos en lista");		
								return;
							}
						}
					}
						#endregion Mas de una coincidencia o ninguna muestra la lista
						#region Encuentra un solo articulo Asigna Codigo y Unidad e idArticulo
					else 
					{
						if (bRastreo) Rastreo("Un codigo de Articulo Encontrado");
						stSelect = string.Format("Exec VentaArticuloEncontrado '{0}', {1}, {2}, '', '', 0, 0, 0", stCod, 0, 2);
						IdArt = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);			
					}
					#endregion Encuentra un solo articulo
					#region Asignaciones
					if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");

					if (IdArt <= 0)
					{
						return;
					}

					// VALIDACIÓN: no permitir duplicados
					if (ExisteArticuloEnGrid(IdArt, e.Cell.Row))
					{
						MessageBox.Show(
							"Este producto ya ha sido ingresado. Utilice la asignación de bodegas para distribuir cantidades.",
							"Point Technology",
							MessageBoxButtons.OK,
							MessageBoxIcon.Warning
							);

						// Limpiar fila
						e.Cell.Row.Cells["idArticulo"].Value = DBNull.Value;
						e.Cell.Row.Cells["Codigo"].Value = "";
						e.Cell.Row.Cells["Articulo"].Value = "";
						e.Cell.Row.Cells["Precio"].Value = 0.00m;

						return;
					}

					stSelect = string.Format("Exec CargaInformArticuloCompra {0}, {1}, 2, '{2}', 1", IdArt, (int)this.txtIdGrupoCliente.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));				
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);
					dr.Read();
					stCod = dr.GetString(0);
					e.Cell.Row.Cells["Articulo"].Value = dr.GetString(1);		
					e.Cell.Row.Cells["idArticulo"].Value = IdArt;
					dPrecio = dr.GetDecimal(3);										
					dr.Close();
					e.Cell.Row.Cells["Codigo"].Value = stCod;	
					e.Cell.Row.Cells["Precio"].Value = dPrecio;

					if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
					{
						string sSQL = string.Format("Exec ValInfoArticulosImportacion {0}, {1}", IdArt, 0);
						string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);

						if (sMensaje.Trim().Length > 0)
						{
							MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						}
					}
					#endregion Asignaciones
				} 
				#endregion Codigo o articulo
						
				#region Cantidad
				if (e.Cell.Column.ToString() == "Cantidad")
				{
					if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value) e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
						
					if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
					{							
						if ((int)e.Cell.Row.Cells["Cantidad"].Value > 1)
						{
							e.Cell.Value = 0;
							MessageBox.Show("Seleccione un Producto antes de Ingresar Cantidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}						
					this.cmdTotal();
				}
				#endregion Cantidad					
				#endregion Fin de Verificacion de Cambio
				#region Precio
					
				if (e.Cell.Column.ToString() == "Precio")
				{
					if (e.Cell.Row.Cells["Precio"].Value == System.DBNull.Value) e.Cell.Row.Cells["Precio"].Value = (decimal)e.Cell.OriginalValue;
					cmdTotal();						
				}
				#endregion Precio						
			}
			finally
			{
				bActualiza = true;
			}
		}

		private bool ExisteArticuloEnGrid(int idArticulo, Infragistics.Win.UltraWinGrid.UltraGridRow filaActual)
		{
			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in ultraGrid1.Rows)
			{
				if (row != filaActual)
				{
					if (row.Cells["idArticulo"].Value != DBNull.Value)
					{
						int id = Convert.ToInt32(row.Cells["idArticulo"].Value);
						if (id == idArticulo)
							return true;
					}
				}
			}
			return false;
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetCompra"].Value = 0;									
				e.Row.Cells["Precio"].Value = 0.00m;
				e.Row.Cells["Subtotal"].Value = 0.00m;			
				e.Row.Cells["Cantidad"].Value = 0;									
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{			
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			#region Decimales para el total
			stFormato = stMoneda + stCeros.Substring(0, 2);
			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["SubTotal"].Format = stFormato; 

			this.txtSubTotal.FormatString = stFormato;
			this.txtIvaTotal.FormatString = stFormato;
			this.txtTotal.FormatString = stFormato;

			this.txtSubTotal.MaskInput = stInput;
			this.txtIvaTotal.MaskInput = stInput;
			this.txtTotal.MaskInput = stInput;
			#endregion
			
			FuncionesProcedimientos.FormatoGrid(e, "Precio", 5);													
			FuncionesProcedimientos.FormatoGrid(e, "SubTotal", 5);			
		}
		private void Rastreo(string stMensaje)
		{
			string stSelect = "Insert into ErrorSalida (Mensaje) Values ('" + stMensaje + "')";
			Funcion.EjecutaSQL(cdsSeteoF, stSelect);
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Codigo")
			{	
				int idDetCompra = 0;
				if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
				int idArticulo = 0;
				if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;			
			}
		}
		public void ConsultaMaestro(int IdOrdenCompra)
		{
			try
			{
				this.btnNuevo.Enabled = true;
				this.btnConsultar.Enabled = true;			
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = false;							
				this.btnAnular.Enabled = false;				
				
				string sSQL = string.Format("Exec ConsultaOrdenDeCompraMaestro {0}", IdOrdenCompra);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);

				dr.Read();

				if (dr.HasRows)
				{
					this.txtIdOrdenCompra.Value = IdOrdenCompra;
					this.txtIdCliente.Value = dr.GetInt32(1);
					this.txtIdGrupoCliente.Value = dr.GetInt32(2);

					this.txtRuc.Text = dr.GetString(3);
					this.txtNombre.Text = dr.GetString(4);

					this.dtFecha.Value = dr.GetDateTime(5);

					this.txtNumero.Text = dr.GetString(6);
					this.txtNotas.Text = dr.GetValue(7).ToString();

					this.txtSubTotal.Value = dr.GetDecimal(8);
					this.txtIvaTotal.Value = dr.GetDecimal(9);
					this.txtTotal.Value = dr.GetDecimal(10);

					this.txtEstado.Value = dr.GetInt32(11);

					bAnulado = dr.GetBoolean(12);

					this.lblUsuario.Text = dr.GetString(13);

					this.chkConsignacion.Checked =
						dr.IsDBNull(20) ? false : dr.GetBoolean(20);
				}

				dr.Close();
			
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleOrdenDeCompra {0}", IdOrdenCompra));
			
				int est = (this.txtEstado.Value == null || this.txtEstado.Value == DBNull.Value)
					? 0
					: Convert.ToInt32(this.txtEstado.Value);

				if (bAnulado || est == 7)
				{
					this.lblEstado.Text = "ANULADO";
				}
				else
				{
					if (est == 1)  this.lblEstado.Text = "PENDIENTE";
					if (est == 3)  this.lblEstado.Text = "PROCESADO";
					if (est == 9 || est == 13) this.lblEstado.Text = "FACTURADO";
					if (est == 5)  this.lblEstado.Text = "APROBADO";
					if (est == 4)  this.lblEstado.Text = "RECHAZADO";
					if (est == 6)  this.lblEstado.Text = "REVISION";
					if (est == 12) this.lblEstado.Text = "FACTURADO SIN MERCADERIA";
				}
				if (!bAnulado)
				{

					if ((int)this.txtEstado.Value == 1) this.btnEditar.Enabled = true;
					else if ((int)this.txtEstado.Value == 3 && miAcceso.BEditarProcesado) this.btnEditar.Enabled = true;
					else if ((int)this.txtEstado.Value == 5 && miAcceso.Editar) this.btnEditar.Enabled = true;

					if (!bAnulado)
					{
						if ((int)this.txtEstado.Value == 1) this.btnAnular.Enabled = true;
						else if ((int)this.txtEstado.Value == 3)this.btnAnular.Enabled = true;
					}					
				}		

				this.btnCancelar.Enabled = true;		
				CambiarEstadoFormulario(
					EstadoFormulario.Consulta);
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}	
		}
		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.ConsultaMaestro((int)this.txtIdOrdenCompra.Value);
			
			if (bAnulado)
			{
				MessageBox.Show("No puede editar registro ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if ((int)this.txtEstado.Value == 3 )
			{
				MessageBox.Show("No puede editar una orden de compra PROCESADA, Consulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}

			if ((int)this.txtEstado.Value == 9 || (int)this.txtEstado.Value == 13 || (int)this.txtEstado.Value == 12)
			{
				MessageBox.Show("No Puede Editar una orden de compra FACTURADA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}
					
			dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
			
			#region Controles
			this.txtRuc.Enabled = true;

			this.dtFecha.Enabled = true;						
			
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-5);
						
			this.txtNotas.Enabled = true;
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;			
			this.btnAnular.Enabled = true;
			this.btnCancelar.Enabled = true;		

			bNuevo = false;
			bEdicion = true;
			#endregion Controles			

			CambiarEstadoFormulario(
				EstadoFormulario.Edicion);
		}
		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			this.cmdTotal();
		}
		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if (this.txtIdOrdenCompra.Value == null || (int)this.txtIdOrdenCompra.Value <= 0)
			{
				MessageBox.Show("No existe una Orden de Compra seleccionada.","Anulación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			int idOrden = (int)this.txtIdOrdenCompra.Value;

			int estadoActual = Funcion.iEscalarSQL(cdsSeteoF,string.Format("SELECT Estado FROM OrdenDeCompra WHERE IdOrdenDeCompra = {0}", idOrden));
			if (estadoActual == 5)
			{
				MessageBox.Show("La Orden de Compra ya se encuentra ANULADA.","Anulación", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (estadoActual == 2 || estadoActual == 3)
			{
				MessageBox.Show("La Orden de Compra ya fue procesada y no puede ser anulada.","Anulación", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (MessageBox.Show("¿Está seguro de anular esta Orden de Compra?", "Confirmar Anulación", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
			{				
				return;
			}

			string sSQL = string.Format("UPDATE OrdenDeCompra SET Estado = 7 WHERE IdOrdenDeCompra = {0}", idOrden);

			if (Funcion.EjecutaSQL(cdsSeteoF, sSQL))
			{
				// =========================
				// NUEVO: ANULAR BODEGAS
				// =========================
				try
				{
					DetalleBodegaService service = new DetalleBodegaService();
					service.AnularBodegasPorOrden(cdsSeteoF, idOrden);
				}
				catch
				{
					// No rompemos flujo actual
				}

				MessageBox.Show("La Orden de Compra fue anulada correctamente.","Anulación", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.ConsultaMaestro((int)this.txtIdOrdenCompra.Value);	

				CambiarEstadoFormulario(
					EstadoFormulario.Consulta);

			}
			else
			{
				MessageBox.Show("Error al intentar anular la Orden de Compra.","Anulación", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void txtRuc_Validating(
			object sender,
			System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion && !procesandoImportacion)
			{
				if (!Validacion.vbClienteExiste(
					this.txtRuc,
					true,
					2,
					6,
					cdsSeteoF))
				{
					this.txtRuc.Focus();

					e.Cancel = true;

					return;
				}

				if (
					Funcion.iEscalarSQL(
					cdsSeteoF,
					string.Format(
					"Exec ValidaGrupoProveedor '{0}'",
					this.txtRuc.Text.ToString()
					)
					) == 0
					)
				{
					MessageBox.Show(
						"En la Ficha Proveedores Seleccione el Grupo al cual Pertenece el Proveedor",
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Stop);

					this.txtRuc.Focus();

					e.Cancel = true;

					return;
				}
			}
		}

		private bool ValidaCantidadesUIOGYE()
		{
			if (ultraGrid1.ActiveCell != null && ultraGrid1.ActiveCell.IsInEditMode)
				ultraGrid1.PerformAction(UltraGridAction.ExitEditMode);
			ultraGrid1.UpdateData();

			foreach (UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.IsAddRow) continue;
				if (dr.Cells.Exists("idArticulo") && dr.Cells["idArticulo"].Value == System.DBNull.Value) continue;
			}
			return true;
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if (!bEliminaAlValidar)
				{
					if (!bNuevo)
					{
						e.DisplayPromptMsg = false;

						string sCodigos = "";
						int i;

						for (i = 0; i < e.Rows.Length; i++)
						{
							if (e.Rows[i].Cells.Exists("Codigo") && e.Rows[i].Cells["Codigo"].Value != DBNull.Value)
							{
								string cod = e.Rows[i].Cells["Codigo"].Value.ToString().Trim();
								if (cod.Length > 0)
								{
									if (sCodigos.Length > 0) sCodigos += ", ";
									sCodigos += cod;
								}
							}
						}

						string msg = "¿Desea borrar definitivamente los ítems seleccionados";
						if (sCodigos.Length > 0) msg += " (" + sCodigos + ")";
						msg += "?\n\nEste proceso NO puede ser revertido.";

						if (MessageBox.Show(msg, "Point Technology",
							MessageBoxButtons.YesNo,
							MessageBoxIcon.Warning,
							MessageBoxDefaultButton.Button2) != DialogResult.Yes)
						{
							e.Cancel = true;
							return;
						}

						// ARRAYLIST (compatible .NET 1.1)
						System.Collections.ArrayList ids = new System.Collections.ArrayList();

						for (i = 0; i < e.Rows.Length; i++)
						{
							if (e.Rows[i].Cells["idDetCompra"].Value != DBNull.Value)
							{
								ids.Add(Convert.ToInt32(e.Rows[i].Cells["idDetCompra"].Value));
							}
						}

						try
						{
							// LLAMADA AL SERVICE
							Latinium.Services.OrdenDeCompra.Bodega.OrdenDeCompraService service =
								new Latinium.Services.OrdenDeCompra.Bodega.OrdenDeCompraService(cdsSeteoF);

							service.EliminarDetalles(ids);

							e.Cancel = false;
						}
						catch (Exception ex)
						{
							e.Cancel = true;
							MessageBox.Show(ex.Message, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
					}
					else
					{
						e.DisplayPromptMsg = false;
					}
				}
				else
				{
					e.DisplayPromptMsg = false;
				}
			}
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if ((int)this.txtIdCliente.Value > 0)
			{
				if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
				{									
					this.txtIdComprobante.Value = 12;
				}
				else
				{			
					this.txtIdComprobante.Value = 1;
				}
			}
		}

		private void frmOrdenDeCompra_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.F5) 
				{				
					this.ConsultaMaestro((int)this.txtIdOrdenCompra.Value);
				}
				if (e.KeyCode == Keys.Escape) this.Close();					
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 				
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
						this.btnCancelar_Click(sender, e);					
			}
		}

//		#region Estado Formulario

//		private void CambiarEstadoFormulario(
//			EstadoFormulario estado)
//		{
//			estadoFormulario = estado;
//
//			bool editando =
//				estado == EstadoFormulario.Nuevo ||
//				estado == EstadoFormulario.Edicion;
//
//			bool procesando =
//				estado == EstadoFormulario.Procesando;
//
//			// =========================
//			// BOTONES PRINCIPALES
//			// =========================
//
//			btnNuevo.Enabled =
//				!procesando &&
//				!editando;
//
//			btnConsultar.Enabled =
//				!procesando &&
//				!editando;
//
//			btnEditar.Enabled =
//				!procesando &&
//				estado == EstadoFormulario.Consulta;
//
//			btnGuardar.Enabled =
//				!procesando &&
//				editando;
//
//			btnCancelar.Enabled =
//				!procesando &&
//				editando;
//
//			btnSalir.Enabled =
//				!procesando;
//
//			// =========================
//			// IMPORTACIÓN
//			// =========================
//
//			btnImportar.Enabled =
//				!procesando &&
//				editando;
//
//			btnPlantilla.Enabled =
//				!procesando &&
//				editando;
//
//			// =========================
//			// ANULAR
//			// =========================
//
//			bool puedeAnular = false;
//
//			if (
//				!procesando &&
//				estado == EstadoFormulario.Consulta
//				)
//			{
//				if (this.txtIdOrdenCompra.Value != null)
//				{
//					int idOrden =
//						(int)this.txtIdOrdenCompra.Value;
//
//					int estadoOC =
//						Funcion.iEscalarSQL(
//						cdsSeteoF,
//						string.Format(
//						"SELECT Estado FROM OrdenDeCompra WHERE IdOrdenDeCompra = {0}",
//						idOrden
//						)
//						);
//
//					// =========================
//					// VALIDAR ESTADO REAL
//					// =========================
//					// 7 = ANULADA
//					// 2 y 3 = PROCESADA
//					// =========================
//
//					puedeAnular =
//						estadoOC != 7 &&
//						estadoOC != 2 &&
//						estadoOC != 3;
//				}
//			}
//
//			btnAnular.Enabled = puedeAnular;
//
//			// =========================
//			// CURSOR
//			// =========================
//
//			this.Cursor =
//				procesando
//				? Cursors.WaitCursor
//				: Cursors.Default;
//		}
//
//		#endregion

		private void CambiarEstadoFormulario(
			EstadoFormulario estado)
		{
			estadoFormulario = estado;

			switch (estado)
			{
				case EstadoFormulario.Inicial:

					ConfigurarEstadoInicial();
					break;

				case EstadoFormulario.Consulta:

					ConfigurarEstadoConsulta();
					break;

				case EstadoFormulario.Nuevo:

					ConfigurarEstadoNuevo();
					break;

				case EstadoFormulario.Edicion:

					ConfigurarEstadoEdicion();
					break;

				case EstadoFormulario.Procesando:

					ConfigurarEstadoProcesando();
					break;
			}
		}

		private void ConfigurarEstadoInicial()
		{
			btnNuevo.Enabled = true;
			btnConsultar.Enabled = true;
			btnEditar.Enabled = false;
			btnGuardar.Enabled = false;
			btnAnular.Enabled = false;
			btnImportar.Enabled = false;
			btnPlantilla.Enabled = false;
			btnCancelar.Enabled = false;
			btnSalir.Enabled = true;

			this.Cursor = Cursors.Default;
		}

		private void ConfigurarEstadoConsulta()
		{
			btnNuevo.Enabled = true;
			btnConsultar.Enabled = true;

			btnEditar.Enabled = false;

			btnGuardar.Enabled = false;
			btnAnular.Enabled =
				PuedeAnularOrden();

			btnImportar.Enabled = false;
			btnPlantilla.Enabled = false;

			btnCancelar.Enabled = true;

			btnSalir.Enabled = true;

			this.Cursor = Cursors.Default;
		}


		private void ConfigurarEstadoNuevo()
		{
			procesandoImportacion = true;
			btnNuevo.Enabled = false;
			btnConsultar.Enabled = false;

			btnEditar.Enabled = false;

			btnGuardar.Enabled = true;

			btnImportar.Enabled = true;
			btnPlantilla.Enabled = true;

			btnCancelar.Enabled = true;

			btnSalir.Enabled = true;

			btnAnular.Enabled = false;

			txtRuc.Enabled = false;

			this.txtNotas.Enabled = false;

			this.Cursor = Cursors.Default;
		}

		private void ConfigurarEstadoEdicion()
		{
			btnNuevo.Enabled = false;
			btnConsultar.Enabled = false;

			btnEditar.Enabled = false;

			btnGuardar.Enabled = true;

			btnImportar.Enabled = true;
			btnPlantilla.Enabled = true;

			btnCancelar.Enabled = true;

			btnSalir.Enabled = true;

			// =========================
			// ANULAR
			// =========================

			btnAnular.Enabled = PuedeAnularOrden();

			this.Cursor = Cursors.Default;
		}

		private void ConfigurarEstadoProcesando()
		{
			btnNuevo.Enabled = false;
			btnConsultar.Enabled = false;

			btnEditar.Enabled = false;
			btnGuardar.Enabled = false;

			btnImportar.Enabled = false;
			btnPlantilla.Enabled = false;

			btnAnular.Enabled = false;

			btnCancelar.Enabled = false;

			btnSalir.Enabled = false;

			this.Cursor = Cursors.WaitCursor;
		}

		private bool PuedeAnularOrden()
		{
			// =========================
			// VALIDAR CONTROL
			// =========================
			if (this.txtIdOrdenCompra.Value == null)
				return false;

			if ((int)this.txtIdOrdenCompra.Value == 0)
				return false;

			int idOrden = 0;

			try
			{
				idOrden =
					Convert.ToInt32(
					this.txtIdOrdenCompra.Value
					);
			}
			catch
			{
				return false;
			}

			// =========================
			// NUEVA ORDEN
			// TODAVÍA NO GUARDADA
			// =========================
			if (idOrden <= 0)
				return true;

			// =========================
			// VALIDAR EXISTENCIA
			// =========================
			int existe =
				Funcion.iEscalarSQL(
				cdsSeteoF,
				string.Format(
				"SELECT COUNT(*) FROM OrdenDeCompra WHERE IdOrdenDeCompra = {0}",
				idOrden
				)
				);

			if (existe <= 0)
				return true;

			// =========================
			// ESTADO
			// =========================
			int estadoOC =
				Funcion.iEscalarSQL(
				cdsSeteoF,
				string.Format(
				"SELECT ISNULL(Estado,0) " +
				"FROM OrdenDeCompra " +
				"WHERE IdOrdenDeCompra = {0}",
				idOrden
				)
				);

			// 7 = ANULADA
			// 2 = PROCESADA
			// 3 = CERRADA

			return
				estadoOC != 7 &&
				estadoOC != 2 &&
				estadoOC != 3;
		}

		#region Orden De Compra Dinámico

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			try
			{
				if (e.Cell == null) return;

				if (e.Cell.Column.Key == "btnBodega")
				{
					UltraGridRow row = e.Cell.Row;
					AbrirBodegas(row);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Error");
			}
		}

		private void AbrirBodegas(UltraGridRow row)
		{
			try
			{
				if (!ValidarRow(row)) return;

				DetalleBodegaService service = new DetalleBodegaService();

				BodegaRequest request = ConstruirRequest(row);

				BodegaResponse response = service.PrepararBodegas(request);

				frmOrdenDeCompraBodega frm = CrearFormulario(response);

				if (frm.ShowDialog() == DialogResult.OK)
				{
					if (response.SoloLectura) return;

					AplicarResultado(row, frm.Resultado, service);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Error");
			}
		}

		private bool ValidarRow(UltraGridRow row)
		{
			if (row == null) return false;

			if (row.Cells["idArticulo"].Value == DBNull.Value)
				return false;

			return true;
		}

		private BodegaRequest ConstruirRequest(UltraGridRow row)
		{
			BodegaRequest request = new BodegaRequest();

			// modo
			request.SoloLectura = true; // !(bNuevo || bEdicion);

			// cantidad
			if (row.Cells["Cantidad"].Value != DBNull.Value)
				request.Cantidad = Convert.ToInt32(row.Cells["Cantidad"].Value);

			// id detalle
			if (row.Cells["idDetCompra"].Value != DBNull.Value)
				request.IdDetOrdenCompra = Convert.ToInt32(row.Cells["idDetCompra"].Value);

			// articulo
			string codigo = "";
			string nombre = "";

			if (row.Cells["Codigo"].Value != DBNull.Value)
				codigo = row.Cells["Codigo"].Value.ToString();

			if (row.Cells["Articulo"].Value != DBNull.Value)
				nombre = row.Cells["Articulo"].Value.ToString();

			request.Articulo = codigo + " - " + nombre;

			// temporales
			request.Temporales = row.Tag as ArrayList;

			// dataset
			request.CdsSeteoF = this.cdsSeteoF;

			return request;
		}

		private frmOrdenDeCompraBodega CrearFormulario(BodegaResponse response)
		{
			frmOrdenDeCompraBodega frm = new frmOrdenDeCompraBodega();

			frm.CantidadTotal = response.CantidadTotal;
			frm.SoloLectura = response.SoloLectura;
			frm.Articulo = response.Articulo;
			frm.DataBodegas = response.DataBodegas;

			return frm;
		}

		private void AplicarResultado(
			UltraGridRow row,
			ArrayList resultado,
			DetalleBodegaService service)
		{
			row.Tag = resultado;

			int total = service.CalcularTotal(resultado);

			row.Cells["Cantidad"].Value = total;
			row.Update();

			ultraGrid1.Refresh();
		}
		#endregion Orden De Compra Dinámico

		private void btnPlantilla_Click(
			object sender,
			System.EventArgs e)
		{
			if (
				estadoFormulario != EstadoFormulario.Nuevo &&
				estadoFormulario != EstadoFormulario.Edicion
				)
			{
				MessageBox.Show(
					"Solo puede descargar plantilla en modo edición",
					"Plantilla",
					MessageBoxButtons.OK,
					MessageBoxIcon.Warning);

				return;
			}

			try
			{
				// =========================
				// BLOQUEA VALIDACIONES
				// =========================

				procesandoImportacion = true;

				CambiarEstadoFormulario(
					EstadoFormulario.Procesando);

				// =========================
				// CONFIGURACION API
				// =========================

				ConfiguracionService config =
					new ConfiguracionService(
					cdsSeteoF);

				string urlBase =
					config.ObtenerUrlApi();

				if (
					urlBase == null ||
					urlBase.Trim() == ""
					)
				{
					MessageBox.Show(
						"No existe configuración de API",
						"Configuración",
						MessageBoxButtons.OK,
						MessageBoxIcon.Warning);

					return;
				}

				string url =
					urlBase +
					"/api/v1/orden-compra/plantilla?idUsuario=1";

				// =========================
				// DESCARGAR ARCHIVO
				// =========================

				System.Net.WebClient client =
					new System.Net.WebClient();

				byte[] data =
					client.DownloadData(
					url);

				using (
					System.IO.MemoryStream ms =
						new System.IO.MemoryStream(data)
						)
				{
					string ruta =
						System.IO.Path.Combine(
						System.IO.Path.GetTempPath(),
						"OrdenCompra.xlsx"
						);

					System.IO.FileStream fs =
						new System.IO.FileStream(
						ruta,
						System.IO.FileMode.Create,
						System.IO.FileAccess.Write
						);

					byte[] buffer =
						ms.ToArray();

					fs.Write(
						buffer,
						0,
						buffer.Length);

					fs.Close();

					System.Diagnostics.Process.Start(
						ruta);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					"Error: " + ex.Message);
			}
			finally
			{
				// =========================
				// RESTAURA VALIDACIONES
				// =========================

				procesandoImportacion = false;

				if (bEdicion)
				{
					CambiarEstadoFormulario(
						EstadoFormulario.Edicion);
				}
				else
				{
					CambiarEstadoFormulario(
						EstadoFormulario.Nuevo);
				}
			}
		}

		private void btnImportar_Click(
			object sender,
			System.EventArgs e)
		{
//			if (procesandoImportacion)
//				return;

			if (
				estadoFormulario != EstadoFormulario.Nuevo &&
				estadoFormulario != EstadoFormulario.Edicion
				)
			{
				MessageBox.Show(
					"Solo puede importar en modo Nuevo o Edición",
					"Importación",
					MessageBoxButtons.OK,
					MessageBoxIcon.Warning);

				return;
			}

			OpenFileDialog open =
				new OpenFileDialog();

			open.Filter =
				"Archivos Excel|*.xls;*.xlsx";

			if (open.ShowDialog() != DialogResult.OK)
				return;

			// =========================
			// GUARDAR ESTADO ORIGINAL
			// =========================

			bool eraEdicion = bEdicion;

			try
			{
				procesandoImportacion = true;

				CambiarEstadoFormulario(
					EstadoFormulario.Procesando);

				// =========================
				// LEER EXCEL
				// =========================

				FileStream fs =
					new FileStream(
					open.FileName,
					FileMode.Open,
					FileAccess.Read);

				byte[] fileBytes =
					new byte[fs.Length];

				fs.Read(
					fileBytes,
					0,
					(int)fs.Length);

				fs.Close();

				string base64 =
					Convert.ToBase64String(fileBytes);

				// =========================
				// LLAMAR API
				// =========================

				ImportarApiService apiService =
					new ImportarApiService(cdsSeteoF);

				ImportacionXmlResult result =
					apiService.ProcesarArchivo(
					base64,
					1);

				// =========================
				// VALIDAR RESPUESTA
				// =========================

				if (result == null)
				{
					MessageBox.Show(
						"Sin respuesta del servidor");

					return;
				}

				if (!result.Ok)
				{
					MessageBox.Show(
						result.Mensaje,
						"Error importación",
						MessageBoxButtons.OK,
						MessageBoxIcon.Error);

					return;
				}

				// =========================
				// LIMPIAR FORMULARIO
				// =========================

				this.btnNuevo_Click(null, null);

				// =========================
				// CABECERA
				// =========================

				this.txtRuc.Text =
					result.RucProveedor;

				this.CargaCliente();
				
				VerificarComprobante();

				if (result.Fecha != DateTime.MinValue)
					this.dtFecha.Value =
						result.Fecha;

				this.txtNotas.Text =
					result.Notas;

				// =========================
				// CONSIGNACIÓN (NUEVO)
				// =========================

				if (result.Consignacion != null)
				{
					this.chkConsignacion.Checked =
						result.Consignacion.ToString().ToUpper() == "SI";
				}
				else
				{
					this.chkConsignacion.Checked = false;
				}

				// =========================
				// DETALLE
				// =========================

				this.ultraGrid1.BeginUpdate();

				try
				{
					if (result.Detalles != null)
					{
						IEnumerator detEnum =
							result.Detalles.GetEnumerator();

						while (detEnum.MoveNext())
						{
							ImportacionDetalle item =
								(ImportacionDetalle)
								detEnum.Current;

							UltraGridRow row =
								this.ultraGrid1
								.DisplayLayout
								.Bands[0]
								.AddNew();

							row.Cells["Codigo"].Value =
								item.Codigo;

							this.ultraGrid1_AfterCellUpdate(
								this.ultraGrid1,
								new CellEventArgs(
								row.Cells["Codigo"]
								)
								);

							row.Cells["Cantidad"].Value =
								item.Cantidad;

							row.Cells["Precio"].Value =
								item.Precio;

							row.Cells["SubTotal"].Value =
								item.Subtotal;

							row.Tag =
								item.Bodegas;
						}
					}
				}
				finally
				{
					this.ultraGrid1.EndUpdate();
				}

				// =========================
				// TOTALES
				// =========================

				this.cmdTotal();

				MessageBox.Show(
					"Importación de Excel finalizada con éxito.");
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					"Error crítico: " + ex.Message);
			}
			finally
			{
				procesandoImportacion = false;

				if (eraEdicion)
				{
					CambiarEstadoFormulario(
						EstadoFormulario.Edicion);
				}
				else
				{
					CambiarEstadoFormulario(
						EstadoFormulario.Nuevo);
				}
			}
		}

		private void VerificarComprobante()
		{
			if ((int)this.txtIdCliente.Value > 0)
			{
				if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
				{									
					this.txtIdComprobante.Value = 12;
				}
				else
				{			
					this.txtIdComprobante.Value = 1;
				}
			}
		}
	}
}
