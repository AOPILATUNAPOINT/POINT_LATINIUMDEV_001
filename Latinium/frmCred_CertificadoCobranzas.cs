using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCred_CertificadoCobranzas.
	/// </summary>
	public class frmCred_CertificadoCobranzas : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdentificacion;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstadoAutorizacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAutorizacion;
		private Infragistics.Win.Misc.UltraLabel lblFecha;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
//		private C1DataRow drSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoTotal;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private System.Windows.Forms.GroupBox gbBotones;
		private Acceso miAcceso;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCred_CertificadoCobranzas()
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

		private void UnloadMe()
		{
			this.Close();
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinEditors.EditorButton editorButton1 = new Infragistics.Win.UltraWinEditors.EditorButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCred_CertificadoCobranzas));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Confirmado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDeOperacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Reliquido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ER");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Unidades");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Unidades");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Solicita_bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EntFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Cancelacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Actual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Solicita_bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EntFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Cancelacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Actual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ir", 0);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			this.txtIdentificacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtEstadoAutorizacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAutorizacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblFecha = new Infragistics.Win.Misc.UltraLabel();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtSaldoTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoAutorizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAutorizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			this.SuspendLayout();
			// 
			// txtIdentificacion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdentificacion.Appearance = appearance1;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			editorButton1.Appearance = appearance2;
			editorButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button3D;
			editorButton1.Text = "...";
			this.txtIdentificacion.ButtonsRight.Add(editorButton1);
			this.txtIdentificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdentificacion.Location = new System.Drawing.Point(72, 56);
			this.txtIdentificacion.MaxLength = 10;
			this.txtIdentificacion.Name = "txtIdentificacion";
			this.txtIdentificacion.Size = new System.Drawing.Size(160, 21);
			this.txtIdentificacion.TabIndex = 273;
			this.txtIdentificacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtIdentificacion_KeyDown);
			this.txtIdentificacion.EditorButtonClick += new Infragistics.Win.UltraWinEditors.EditorButtonEventHandler(this.txtIdentificacion_EditorButtonClick);
			this.txtIdentificacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIdentificacion_KeyPress);
			this.txtIdentificacion.Validated += new System.EventHandler(this.txtIdentificacion_Validated);
			this.txtIdentificacion.ValueChanged += new System.EventHandler(this.txtIdentificacion_ValueChanged);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 14F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Red;
			this.lblEstado.Location = new System.Drawing.Point(376, 56);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 282;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(16, 32);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 278;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtEstadoAutorizacion
			// 
			this.txtEstadoAutorizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstadoAutorizacion.Enabled = false;
			this.txtEstadoAutorizacion.Location = new System.Drawing.Point(608, 48);
			this.txtEstadoAutorizacion.Name = "txtEstadoAutorizacion";
			this.txtEstadoAutorizacion.PromptChar = ' ';
			this.txtEstadoAutorizacion.ReadOnly = true;
			this.txtEstadoAutorizacion.Size = new System.Drawing.Size(32, 21);
			this.txtEstadoAutorizacion.TabIndex = 281;
			this.txtEstadoAutorizacion.Visible = false;
			// 
			// txtIdAutorizacion
			// 
			this.txtIdAutorizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAutorizacion.Enabled = false;
			this.txtIdAutorizacion.Location = new System.Drawing.Point(608, 80);
			this.txtIdAutorizacion.Name = "txtIdAutorizacion";
			this.txtIdAutorizacion.PromptChar = ' ';
			this.txtIdAutorizacion.ReadOnly = true;
			this.txtIdAutorizacion.Size = new System.Drawing.Size(32, 21);
			this.txtIdAutorizacion.TabIndex = 280;
			this.txtIdAutorizacion.Visible = false;
			// 
			// lblFecha
			// 
			this.lblFecha.AutoSize = true;
			this.lblFecha.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblFecha.Location = new System.Drawing.Point(16, 8);
			this.lblFecha.Name = "lblFecha";
			this.lblFecha.Size = new System.Drawing.Size(34, 15);
			this.lblFecha.TabIndex = 279;
			this.lblFecha.Text = "Fecha";
			// 
			// dtFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFecha.CausesValidation = false;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 271;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(608, 16);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.ReadOnly = true;
			this.txtIdCliente.Size = new System.Drawing.Size(32, 21);
			this.txtIdCliente.TabIndex = 277;
			this.txtIdCliente.Visible = false;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(16, 80);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(40, 14);
			this.ultraLabel2.TabIndex = 276;
			this.ultraLabel2.Text = "Cliente";
			// 
			// txtNombre
			// 
			appearance4.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance4.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance4;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(72, 80);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(336, 21);
			this.txtNombre.TabIndex = 274;
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(16, 56);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(40, 14);
			this.ultraLabel4.TabIndex = 275;
			this.ultraLabel4.Text = "Cédula";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn3.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn4.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(bool);
			ultraDataColumn9.DefaultValue = false;
			ultraDataColumn11.DataType = typeof(bool);
			ultraDataColumn11.DefaultValue = false;
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn12.DefaultValue = 0;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12});
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
			// cmbBodega
			// 
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance5;
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(72, 32);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(240, 22);
			this.cmbBodega.TabIndex = 284;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance7;
			ultraGridColumn5.Header.Caption = "Cantidad";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 8;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 8;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 8;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 9F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance11.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(576, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(24, 0);
			this.ultraGrid1.TabIndex = 287;
			this.ultraGrid1.Visible = false;
			// 
			// txtSaldoTotal
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoTotal.Appearance = appearance14;
			this.txtSaldoTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoTotal.Enabled = false;
			this.txtSaldoTotal.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSaldoTotal.FormatString = "#,##0.00";
			this.txtSaldoTotal.Location = new System.Drawing.Point(472, 8);
			this.txtSaldoTotal.MinValue = 0;
			this.txtSaldoTotal.Name = "txtSaldoTotal";
			this.txtSaldoTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoTotal.PromptChar = ' ';
			this.txtSaldoTotal.Size = new System.Drawing.Size(80, 22);
			this.txtSaldoTotal.TabIndex = 289;
			this.txtSaldoTotal.Visible = false;
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.AutoSize = true;
			this.ultraLabel7.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel7.Location = new System.Drawing.Point(384, 8);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(76, 15);
			this.ultraLabel7.TabIndex = 288;
			this.ultraLabel7.Text = "SALDO TOTAL";
			this.ultraLabel7.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(System.DateTime);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.DateTime);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(bool);
			ultraDataColumn34.DataType = typeof(System.DateTime);
			ultraDataColumn35.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35});
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource3;
			appearance15.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance15;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.Caption = "Local";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn10.Header.VisiblePosition = 15;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance16;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Width = 90;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Width = 110;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Width = 237;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn16.Header.Caption = "Banco";
			ultraGridColumn16.Header.VisiblePosition = 9;
			ultraGridColumn16.Width = 105;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn17.Header.VisiblePosition = 10;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance17;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 11;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance18;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 12;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn20.Header.VisiblePosition = 13;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn21.Header.VisiblePosition = 14;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn22.Header.VisiblePosition = 16;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn23.Header.VisiblePosition = 17;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 18;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 19;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance19.Image = ((object)(resources.GetObject("appearance19.Image")));
			appearance19.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn27.CellButtonAppearance = appearance19;
			ultraGridColumn27.Header.Caption = "...";
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn27.Width = 30;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn27});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance20;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance21.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance21;
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance23.FontData.BoldAsString = "True";
			appearance23.FontData.Name = "Arial";
			appearance23.FontData.SizeInPoints = 8F;
			appearance23.ForeColor = System.Drawing.Color.White;
			appearance23.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance23;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance24.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance24.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance26;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(0, 120);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(940, 212);
			this.uGridDocumentos.TabIndex = 765;
			this.uGridDocumentos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDocumentos_KeyDown);
			this.uGridDocumentos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_ClickCellButton);
			this.uGridDocumentos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDocumentos_InitializeLayout);
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(-24, 104);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1024, 8);
			this.gbBotones.TabIndex = 766;
			this.gbBotones.TabStop = false;
			// 
			// frmCred_CertificadoCobranzas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(944, 334);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.uGridDocumentos);
			this.Controls.Add(this.txtSaldoTotal);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtIdentificacion);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.txtEstadoAutorizacion);
			this.Controls.Add(this.txtIdAutorizacion);
			this.Controls.Add(this.lblFecha);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.ultraLabel4);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCred_CertificadoCobranzas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "CREDIPOINT - CERTIFICADO";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCred_CertificadoCobranzas_KeyDown);
			this.Load += new System.EventHandler(this.frmCred_CertificadoCobranzas_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoAutorizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAutorizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iBodega = 0;

		private void frmCred_CertificadoCobranzas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083204");
			if (!Funcion.Permiso("1197", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Certificados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);	
//			FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 43, (DateTime)this.dtFecha.Value, cdsSeteoF, true);

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodega > 0) this.cmbBodega.Value = iBodega; 
			else 
			{
				this.cmbBodega.Enabled = true;
				this.cmbBodega.Value = 12;
				this.cmbBodega.Focus ();
			}

			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtIdentificacion_Validated(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un Local")) return;
			string sIdentificacion = this.txtIdentificacion.Text.ToString();

			LimpiaControles(true);
	
			this.txtIdentificacion.Text = sIdentificacion;

			CargaCliente();	
	
			string sNumero = "";
//			int iCont =0;
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drH in this.uGridDocumentos.Rows.All)
//			{
//				sNumero = Convert.ToString( drH.Cells["Numero"].Value).ToString ();
//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drs in this.uGridDocumentos.Rows.All)
//				{
//					if (sNumero == Convert.ToString( drs.Cells["Numero"].Value).ToString () ) 
//					{
//						iCont = iCont +1 ;
//					}
//				}
//			}
//			if(iCont >1)
//			{
//				int idCompra = 0;
//				//         sNumero =  	Convert.ToString( drD.Cells["Numero"].Value).ToString ();
//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drH in this.uGridDocumentos.Rows.All)
//				{
//					if (Convert.ToInt32( (int)drH.Cells["idEntidadFinanciera"].Value) == 4 || Convert.ToInt32( (int)drH.Cells["idEntidadFinanciera"].Value) == 3 )
//					{
//						drH.Cells["Sel"].Value = true;
////						idCompra = (int) drH.Cells["idCompra"].Value ;
//					}
//				}
//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drH in this.uGridDocumentos.Rows.All)
//				{
//					if (Convert.ToInt32( (int)drH.Cells["idEntidadFinanciera"].Value) == 3 && Convert.ToBoolean( drH.Cells["Sel"].Value) == true )
//					{
//						drH.Cells["Sel"].Value = false;
////						drH.Cells["idCompra"].Value = idCompra ;
//					}
//				}

//			}
		}
    
		private void CargaCliente()
		{
			string sSQLValB = string.Format("Exec  Cred_SolicitudCertificado '{0}',{1}", this.txtIdentificacion.Text.ToString (), (int) this.cmbBodega.Value);
			SqlDataReader drBloqueo = Funcion.rEscalarSQL(cdsSeteoF, sSQLValB, true);
			drBloqueo.Read();
			if (drBloqueo.HasRows)
			{

						if (drBloqueo.GetValue(15) != System.DBNull.Value)this.txtIdCliente.Value = drBloqueo.GetInt32(15);
				    if (drBloqueo.GetValue(5) != System.DBNull.Value)this.txtIdentificacion.Text= drBloqueo.GetString(5);
				    if (drBloqueo.GetValue(6) != System.DBNull.Value)this.txtNombre.Text= drBloqueo.GetString(6);
					
			}
			drBloqueo.Close();
				
			if ((int)this.txtIdCliente.Value > 0) 
			{
				ConsultaFacturas(this.txtIdentificacion.Text.ToString());
			}
			else
			{
				this.txtIdCliente.Value = 0;
				this.txtNombre.Text = "";	
				string sSQL = string.Format("Exec [Cred_SolicitudCertificado] '', 0");
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDocumentos);
			}
		}

		private void ConsultaFacturas(string sIdentificacion)
		{
			string sSQL = string.Format("Exec [Cred_SolicitudCertificado] '{0}', {1}", this.txtIdentificacion.Text, (int)this.cmbBodega.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDocumentos);
		}

		private void ConsultaTablaDeArticulos(int idCompra)
		{
			string sSQL = string.Format("Exec DetalleReporteProducto {0}", idCompra);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);		
		}

		private void txtIdentificacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{					
				LimpiaControles(true);

				string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(false, 46);
				if (sIdentificacion.Length > 0)
				{						
					this.txtIdentificacion.Text = sIdentificacion;
					CargaCliente();				
				}				
			}
		}

		private void txtIdentificacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.uGridDocumentos.Focus();
		}

		
		#region Detalle_Producto
		private void Cobranza()
		{
			if (this.uGridDocumentos.ActiveRow != null)
			{		
				try
				{
//					ConsultaTablaDeArticulos();
//					LimpiaControles(false);
											
	
//					string sSQL = string.Format("Exec DetalleReporteProducto {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
//					FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);

//					ConsultaTablaDeArticulos();
					
				}
				catch (Exception exc)
				{
					MessageBox.Show(string.Format("Exception Type: {0} {1}", exc.GetType(), exc.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				}
			}
		}
#endregion Detalle_Producto



		private void txtIdentificacion_EditorButtonClick(object sender, Infragistics.Win.UltraWinEditors.EditorButtonEventArgs e)
		{
			LimpiaControles(true);

			string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(false, 46);

			if (sIdentificacion.Length > 0)
			{						
				this.txtIdentificacion.Text = sIdentificacion;				

				CargaCliente();
			}
		}

		private void LimpiaControles(bool bCliente)
		{
			#region Limpia Informacion
			if (bCliente)
			{				
				this.txtIdCliente.Value = 0;
				this.txtIdentificacion.Text = "";
				this.txtNombre.Text = "";				

			}

		
			#endregion Limpia Informacion
		}

		private void GridFacturas_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
		Cobranza();
		}

		private void GridFacturas_Click(object sender, System.EventArgs e)
		{
			Cobranza();	
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
		
		}

		private void uGridDocumentos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Ir")
			{
				if (e.Cell.Row.Cells["idCompra"].Value != System.DBNull.Value  )
				{
					if ( (bool)e.Cell.Row.Cells["Sel"].Value == true )
					{
						MessageBox.Show("El documento se Encuentra en FONDO DE GARANTIA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						return;
					}
					if ( (bool)e.Cell.Row.Cells["Sel"].Value == true )
					{
						MessageBox.Show("El Credito se encuentra en FONDO DE GARANTIA, comuniquese con el departamento de Cobranzas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						return;
					}
         int icuota = 0;
					if (Convert.ToDecimal (e.Cell.Row.Cells["Saldo"].Value) > 0 )
					{

						string ssql= string.Format ("Exec Consulta_Cbo_GestionesDeCobranzas_Tabla  {0}, '{1}'",(int)e.Cell.Row.Cells["idCompra"].Value , 
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd") ); 
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
						dr.Read();
						if (dr.HasRows)
						{		
					   if (dr.GetValue(0) != System.DBNull.Value)icuota = dr.GetInt32 (0);
						}
			
						dr.Close();

						if (icuota  == 0) 
						{
							MessageBox.Show("El Cliente no tiene Cuotas Pendientes de Pago ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							return;
						}
						if ( (int)e.Cell.Row.Cells["idEntidadFinanciera"].Value == 2 )
						{
							MessageBox.Show("El Certificado de Saldos Pendientes no se puede generar para Solidario ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							return;
						}

						#region Impresion 
						ParameterFields paramFields = new ParameterFields();

						//			ParameterFields paramFields = new ParameterFields ();

						ParameterField paramFieldCedula = new ParameterField ();
						ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
						paramFieldCedula.ParameterFieldName = "@Ruc";
						ValCedula.Value = e.Cell.Row.Cells["Ruc"].Value.ToString() ;
						paramFieldCedula.CurrentValues.Add (ValCedula);
						paramFields.Add (paramFieldCedula);

						ParameterField pfIdBodega= new ParameterField();
						ParameterDiscreteValue ValIdBodega= new ParameterDiscreteValue();
						pfIdBodega.ParameterFieldName = "@bodega";
						ValIdBodega.Value = Convert.ToInt32(this.cmbBodega.Value );
						pfIdBodega.CurrentValues.Add(ValIdBodega);
						paramFields.Add(pfIdBodega);

						ParameterField pfIdOrigen = new ParameterField();
						ParameterDiscreteValue ValIdOrigen = new ParameterDiscreteValue();
						pfIdOrigen.ParameterFieldName = "@idcompra";
						ValIdOrigen.Value = Convert.ToInt32((int)e.Cell.Row.Cells["idCompra"].Value );
						pfIdOrigen.CurrentValues.Add(ValIdOrigen);
						paramFields.Add(pfIdOrigen);


//
////
//						ParameterField pfIdOrigena = new ParameterField();
//						ParameterDiscreteValue ValIdOrigena = new ParameterDiscreteValue();
//						pfIdOrigena.ParameterFieldName = "@idcompra";
//						ValIdOrigena.Value = Convert.ToInt32((int)e.Cell.Row.Cells["idCompra"].Value );
//						pfIdOrigena.CurrentValues.Add(ValIdOrigena);
//						paramFields.Add(pfIdOrigena);
//
//						ParameterField paramFieldFecha = new ParameterField ();
//						ParameterDiscreteValue discreteValFecha = new ParameterDiscreteValue ();
//						paramFieldFecha.ParameterFieldName = "@Fecha";
//						discreteValFecha.Value = this.dtFecha.Value;
//						paramFieldFecha.CurrentValues.Add (discreteValFecha);
//						paramFields.Add (paramFieldFecha);
				
		
						Reporte Reporte = new Reporte("SolicitudCertificadoCuotas.rpt", "");
						Reporte.MdiParent = this.MdiParent;			
						Reporte.crVista.ParameterFieldInfo = paramFields;
						Reporte.Show();
//SolicitudCertificadoCuotasDetalladas

						#endregion Impresion 
//						MessageBox.Show(" NO ES FACTIBLE ENTREGAR EL CERTIFICADO POR ESTAR CON CREDITO EN CURSO. ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//						return;
					}
					else
					{
						if ( (int)e.Cell.Row.Cells["idEntidadFinanciera"].Value == 2 )
						{
							MessageBox.Show("El Certificado es para Clientes Pichincha/Point", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							return;
						}
						#region Impresion 
						ParameterFields paramFields = new ParameterFields();

						//			ParameterFields paramFields = new ParameterFields ();

						ParameterField pfIdOrigen = new ParameterField();
						ParameterDiscreteValue ValIdOrigen = new ParameterDiscreteValue();
						pfIdOrigen.ParameterFieldName = "@idcompra";
						ValIdOrigen.Value = Convert.ToInt32((int)e.Cell.Row.Cells["idCompra"].Value );
						pfIdOrigen.CurrentValues.Add(ValIdOrigen);
						paramFields.Add(pfIdOrigen);
				
						ParameterField paramFieldCedula = new ParameterField ();
						ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
						paramFieldCedula.ParameterFieldName = "@Ruc";
						ValCedula.Value = e.Cell.Row.Cells["Ruc"].Value.ToString() ;
						paramFieldCedula.CurrentValues.Add (ValCedula);
						paramFields.Add (paramFieldCedula);

						ParameterField pfIdBodega= new ParameterField();
						ParameterDiscreteValue ValIdBodega= new ParameterDiscreteValue();
						pfIdBodega.ParameterFieldName = "@bodega";
						ValIdBodega.Value = Convert.ToInt32(this.cmbBodega.Value );
						pfIdBodega.CurrentValues.Add(ValIdBodega);
						paramFields.Add(pfIdBodega);

						Reporte Reporte = new Reporte("SolicitudCertificadoCredito.rpt", "");
						Reporte.MdiParent = this.MdiParent;			
						Reporte.crVista.ParameterFieldInfo = paramFields;
						Reporte.Show();
						#endregion Impresion 
					
					}
				}
//				this.Close ();
			}
		}

		private void frmCred_CertificadoCobranzas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridDocumentos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 9, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
		}

		private void uGridDocumentos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtIdentificacion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

//		private bool ValidaCobro()
//		{
//			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha", true, "Cobros", cdsSeteoF)) return false;
//
//			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return false;
//
//			if (!Validacion.vbIdentificacion(this.txtIdentificacion, 5, cdsSeteoF)) return false;
//
//			if (!Validacion.vbClienteExiste(this.txtIdentificacion, false, 46, 9, cdsSeteoF)) return false;		
//			
//			if (GridFacturas.Rows.Count == 0)
//			{
//				MessageBox.Show("No existen documentos pendientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
//				this.txtIdentificacion.Focus();
//				return false;
//			}
//			if (GridFacturas.ActiveRow == null)
//			{
//				MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);        				
//				this.txtIdentificacion.Focus();
//				return false;
//			}
//			if (this.gridPagos.Rows.Count == 0)
//			{
//				MessageBox.Show("Seleccione un documento para cargar la tabla de amortización", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);        				
//				this.txtIdentificacion.Focus();
//				return false;
//			}
//
//			if (Math.Round(Convert.ToDecimal(this.txtSaldoTotal.Value), 2) == 0.00m)
//			{
//				MessageBox.Show("El Documento esta Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);        				
//				this.txtIdentificacion.Focus();
//				return false;
//			}						
//
//			return true;
//		}	


	}
}
