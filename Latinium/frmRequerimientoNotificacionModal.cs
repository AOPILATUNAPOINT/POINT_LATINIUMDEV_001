using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimientoNotificacionModal.
	/// </summary>
	public class frmRequerimientoNotificacionModal : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoNotificacion;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgRequerimientoNotificacion;
		/// <summary>
		/// Variable del diseñador requerida.
		int Color1 = 0;
		int Color2 = 0;
		int Color3 = 0;
		private System.Windows.Forms.Timer tRequerimientoNotificacion;
		private System.ComponentModel.IContainer components;
		static public bool bAbiertoRequerimientoModal = false;
		public int idRequerimientoRol = 0;
		public string sRequerimientoRol = "";
		public int idRequerimientoDepartamento = 0;
		
		public frmRequerimientoNotificacionModal()
		{
			//Color1 = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "Color1", frmRequerimiento.iConfiguracionColorNotificacion),false);
			//Color2 = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "Color2", frmRequerimiento.iConfiguracionColorNotificacion),false);
			//Color3 = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "Color3", frmRequerimiento.iConfiguracionColorNotificacion),false);
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoNotificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idResponsable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Leido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSeguimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Responsable");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Asignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSeguimientoPrincipal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Respuesta");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoNotificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idResponsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsignado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Leido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeguimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asignado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeguimientoPrincipal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Respuesta", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeguimientoAnadir", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoNotificacionModal));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 1);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.udsRequerimientoNotificacion = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.udgRequerimientoNotificacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.tRequerimientoNotificacion = new System.Windows.Forms.Timer(this.components);
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoNotificacion)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoNotificacion)).BeginInit();
			this.SuspendLayout();
			// 
			// udsRequerimientoNotificacion
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn3.DataType = typeof(System.UInt32);
			ultraDataColumn3.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn6.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn7.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			this.udsRequerimientoNotificacion.Band.Columns.AddRange(new object[] {
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
																																						 ultraDataColumn12,
																																						 ultraDataColumn13});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.udgRequerimientoNotificacion);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(272, 160);
			this.groupBox1.TabIndex = 14;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// udgRequerimientoNotificacion
			// 
			this.udgRequerimientoNotificacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgRequerimientoNotificacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgRequerimientoNotificacion.DataSource = this.udsRequerimientoNotificacion;
			appearance1.BackColor = System.Drawing.Color.White;
			this.udgRequerimientoNotificacion.DisplayLayout.Appearance = appearance1;
			this.udgRequerimientoNotificacion.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.CardSettings.LabelWidth = 7;
			ultraGridBand1.CardSettings.ShowCaption = false;
			ultraGridBand1.CardSettings.Width = 190;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.AutoEdit = false;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.VisiblePosition = 7;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 58;
			ultraGridColumn3.Header.VisiblePosition = 11;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 61;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "";
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 27;
			ultraGridColumn5.AutoEdit = false;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 9;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 114;
			ultraGridColumn6.AutoEdit = false;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 67;
			ultraGridColumn7.Header.VisiblePosition = 10;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 12;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 167;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 169;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn11.Format = "dd/MM/yyyy hh:mm:ss";
			ultraGridColumn11.Header.Caption = "Fecha";
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 788;
			ultraGridColumn12.Header.VisiblePosition = 14;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.AutoEdit = false;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Width = 199;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			appearance2.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance2;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellButtonAppearance = appearance3;
			ultraGridColumn14.Header.Caption = "...";
			ultraGridColumn14.Header.Fixed = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn14.Width = 20;
			ultraGridColumn15.Header.Caption = "N°";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 278;
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
																										 ultraGridColumn15});
			ultraGridBand1.Header.Caption = "";
			this.udgRequerimientoNotificacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.udgRequerimientoNotificacion.DisplayLayout.GroupByBox.Hidden = true;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.udgRequerimientoNotificacion.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgRequerimientoNotificacion.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoNotificacion.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoNotificacion.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.udgRequerimientoNotificacion.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgRequerimientoNotificacion.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.udgRequerimientoNotificacion.Location = new System.Drawing.Point(16, 16);
			this.udgRequerimientoNotificacion.Name = "udgRequerimientoNotificacion";
			this.udgRequerimientoNotificacion.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.udgRequerimientoNotificacion.Size = new System.Drawing.Size(240, 128);
			this.udgRequerimientoNotificacion.TabIndex = 14;
			this.udgRequerimientoNotificacion.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.udgRequerimientoNotificacion_InitializeRow);
			this.udgRequerimientoNotificacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.udgRequerimientoNotificacion_ClickCellButton);
			this.udgRequerimientoNotificacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.udgRequerimientoNotificacion_InitializeLayout);
			// 
			// tRequerimientoNotificacion
			// 
			this.tRequerimientoNotificacion.Enabled = true;
			this.tRequerimientoNotificacion.Interval = 1000;
			this.tRequerimientoNotificacion.Tick += new System.EventHandler(this.tRequerimientoNotificacion_Tick);
			// 
			// frmRequerimientoNotificacionModal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(288, 182);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoNotificacionModal";
			this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
			this.Text = "Notificaciones";
			this.Load += new System.EventHandler(this.frmRequerimientoNotificacionModal_Load);
			this.Closed += new System.EventHandler(this.frmRequerimientoNotificacionModal_Closed);
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoNotificacion)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoNotificacion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoNotificacionModal_Load(object sender, System.EventArgs e)
		{
			//Mostrar notificaicones
			RequerimientoNotificacionCon();
			//
			LocationModal();
			//Intervalo de refresh al timer parametrizado
			this.tRequerimientoNotificacion.Interval = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "Valor", frmRequerimiento.iConfiguracionTimerbtnRequerimiento),false);
			//Muestra que se encuentra abierto 
			bAbiertoRequerimientoModal = true;
		}

		private void LocationModal()
		{
			//Localización del notificador
			int Valiorx = 0;
			int Valiory = 0;
			if(MenuLatinium.getbAbreVentaGeneral)
			{
				Valiorx = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "TamanoMinimo", frmRequerimiento.iConfiguracionAudioNotificacionPosicion),false);
				Valiory = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "TamanoMaximo", frmRequerimiento.iConfiguracionAudioNotificacionPosicion),false);
			}
			else
			{
				Valiorx = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "TamanoMinimo", frmRequerimiento.iConfiguracionAudioNotificacionPosicionSinResumen),false);
				Valiory = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "TamanoMaximo", frmRequerimiento.iConfiguracionAudioNotificacionPosicionSinResumen),false);
			}
			Screen scr = Screen.FromPoint(Location);
			this.Location = new Point (scr.WorkingArea.Right - Width - Valiorx, scr.WorkingArea.Bottom - Height - Valiory);
		}

		private void RequerimientoNotificacionLeido(int idRequerimientoNotificacion)
		{
			string squery = "";
			try
			{
				squery = String.Format("Exec RequerimientoNotificacionLeido {0}", idRequerimientoNotificacion);
				Funcion.EjecusionEjecutaSQL(frmRequerimiento.sconexionPoint, squery, false);
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
			}
		}

		private void RequerimientoNotificacionCon()
		{
			string squery = "";
			try
			{
				squery = String.Format("Exec RequerimientoNotificacionCon {0}, {1}", MenuLatinium.idUsuarioRequerimiento, 0);
				this.udgRequerimientoNotificacion.DataSource = FuncionesProcedimientos.dtEjecucionGeneral(frmRequerimiento.sconexionPoint, squery);
			  this.udgRequerimientoNotificacion = Funcion.ContadorFilas(this.udgRequerimientoNotificacion, "Row");
			}
			catch(Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionPoint, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
			}
		}

		private void udgRequerimientoNotificacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void udgRequerimientoNotificacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int idResponsable = 0;
			int idRequerimientoTipo = 0;
			int idEmpresa = 0;
			int idAsignado = 0;
			string Titulo = "";
			string Descripcion = "";
			DateTime FechaAsignacion;
			DateTime FechaTentativaEntrega;
			string NotaInterna = "";
			int idEstado = 0;
			int idPantalla = 0;
			int idSeguimiento = 0;
			int idRequerimiento = 0;
			string Numero = "";
			int idRequerimientoDepartamento = 0;

			if (e.Cell.Column.ToString() == "btnSeguimientoAnadir")	
			{
				try
				{
					idRequerimiento = (int)e.Cell.Row.Cells["idRequerimiento"].Value;
					idResponsable = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","idPersonal",idRequerimiento), false);
					idRequerimientoTipo = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","idRequerimientoTipo",idRequerimiento), false);
					idEmpresa = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","idEmpresa",idRequerimiento), false);
					idSeguimiento = (int)e.Cell.Row.Cells["idSeguimientoPrincipal"].Value;
					idAsignado = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC SeguimientoModel '{0}',{1}","idPersonal",idSeguimiento), false);
					Titulo = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","Titulo",idRequerimiento));
					Descripcion = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","Descripcion",idRequerimiento));
					FechaAsignacion = Funcion.dtEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","FechaAsignacion",idRequerimiento));
					FechaTentativaEntrega = Funcion.dtEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","FechaTentativaEntrega",idRequerimiento));
					NotaInterna = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","NotaInterna",idRequerimiento));
					idEstado = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","idEstado",idRequerimiento), false);
					idPantalla = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","idPantalla",idRequerimiento), false);
					Numero = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","Numero",idRequerimiento));
					idRequerimientoDepartamento = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoModel '{0}',{1}","idRequerimientoDepartamento",idRequerimiento), false);

					using (frmSeguimientoAnadir frmS = new frmSeguimientoAnadir(
									 idResponsable,
									 idRequerimientoTipo,
									 idEmpresa,
									 idAsignado,
									 Titulo,
									 Descripcion,
									 FechaAsignacion,
									 FechaTentativaEntrega,
									 NotaInterna,
									 idEstado,
									 idPantalla,
									 idSeguimiento,
									 idRequerimiento,
									 Numero,
									 idRequerimientoDepartamento))
					{		
						if (DialogResult.OK == frmS.ShowDialog())
						{
					 
						}
						//Cambiar la notificación a leído
						int idRequerimientoNotificacion = (int)e.Cell.Row.Cells["idRequerimientoNotificacion"].Value;
						RequerimientoNotificacionLeido(idRequerimientoNotificacion);
						//Enlistar las notificaciones
						RequerimientoNotificacionCon();
						//Disminuir 1 al leerlo
						if(frmRequerimientoNotificacionAlerta.iContadorNotificaciones> 0)
						{
							frmRequerimientoNotificacionAlerta.iContadorNotificaciones = frmRequerimientoNotificacionAlerta.iContadorNotificaciones - 1;
						}
					}
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
					return;
				}
			}
		}

		private void udgRequerimientoNotificacion_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
//			if ((bool)e.Row.Cells["Leido"].Value == false)
//			{
//				e.Row.Appearance.BackColor = Color.FromArgb(Color1, Color2, Color3);
//			}
		}

		private void tRequerimientoNotificacion_Tick(object sender, System.EventArgs e)
		{
			LocationModal();
			RequerimientoNotificacionCon();
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void frmRequerimientoNotificacionModal_Closed(object sender, System.EventArgs e)
		{
			bAbiertoRequerimientoModal = false;
		}
	}
}
