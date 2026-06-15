using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

using System.IO;
using DevExpress.XtraEditors;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCred_ReporteVerificacionTelefonica.
	/// </summary>
	public class frmCred_ReporteVerificacionTelefonica : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEIdentificacion;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNENombres;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_Solicitud;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid dgVerificacion;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		bool bCarga = false;

		public frmCred_ReporteVerificacionTelefonica()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		string scedula ="";
		bool bpantalla = false;
		public frmCred_ReporteVerificacionTelefonica(string SCedula, bool BPantalla)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			scedula = SCedula;
			bpantalla =  BPantalla;
      
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
		private Acceso miAcceso;

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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_VerificacionTelefonica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contacto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Referencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_VerificacionTelefonica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Referencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.txtNEIdentificacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNENombres = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtIdCre_Solicitud = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dgVerificacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.txtNEIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNENombres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Solicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dgVerificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			this.SuspendLayout();
			// 
			// txtNEIdentificacion
			// 
			this.txtNEIdentificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEIdentificacion.Location = new System.Drawing.Point(8, 24);
			this.txtNEIdentificacion.Name = "txtNEIdentificacion";
			this.txtNEIdentificacion.Size = new System.Drawing.Size(120, 21);
			this.txtNEIdentificacion.TabIndex = 127;
			this.txtNEIdentificacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEIdentificacion_KeyDown);
			this.txtNEIdentificacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNEIdentificacion_KeyPress);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(93, 16);
			this.label1.TabIndex = 349;
			this.label1.Text = "Ingrese la Cédula";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.Location = new System.Drawing.Point(136, 24);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(80, 23);
			this.btnVer.TabIndex = 352;
			this.btnVer.Text = "Ver";
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(111, 16);
			this.label2.TabIndex = 353;
			this.label2.Text = "Apellidos y Nombres ";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNENombres
			// 
			this.txtNENombres.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNENombres.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNENombres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNENombres.Enabled = false;
			this.txtNENombres.Location = new System.Drawing.Point(8, 72);
			this.txtNENombres.Name = "txtNENombres";
			this.txtNENombres.Size = new System.Drawing.Size(360, 19);
			this.txtNENombres.TabIndex = 354;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.Location = new System.Drawing.Point(1104, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(80, 23);
			this.btnImprimir.TabIndex = 655;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.Location = new System.Drawing.Point(1104, 40);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 656;
			this.btnSalir.Text = "Salir";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtIdCre_Solicitud
			// 
			this.txtIdCre_Solicitud.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCre_Solicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_Solicitud.Enabled = false;
			this.txtIdCre_Solicitud.Location = new System.Drawing.Point(1136, 72);
			this.txtIdCre_Solicitud.Name = "txtIdCre_Solicitud";
			this.txtIdCre_Solicitud.PromptChar = ' ';
			this.txtIdCre_Solicitud.Size = new System.Drawing.Size(48, 21);
			this.txtIdCre_Solicitud.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_Solicitud.TabIndex = 762;
			this.txtIdCre_Solicitud.Visible = false;
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
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(bool);
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
																																 ultraDataColumn11});
			// 
			// dgVerificacion
			// 
			this.dgVerificacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.dgVerificacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.dgVerificacion.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.dgVerificacion.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 150;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 150;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 90;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 90;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 180;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 150;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 150;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 350;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 22;
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
																										 ultraGridColumn11});
			this.dgVerificacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.dgVerificacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.dgVerificacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.dgVerificacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.dgVerificacion.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 8F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.dgVerificacion.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgVerificacion.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgVerificacion.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.dgVerificacion.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dgVerificacion.Location = new System.Drawing.Point(8, 104);
			this.dgVerificacion.Name = "dgVerificacion";
			this.dgVerificacion.Size = new System.Drawing.Size(1176, 248);
			this.dgVerificacion.TabIndex = 763;
			this.dgVerificacion.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.dgVerificacion_CellChange_1);
			// 
			// optEstados
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance6;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance7;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Cliente";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Garante";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(224, 24);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(144, 24);
			this.optEstados.TabIndex = 764;
			this.optEstados.Text = "Cliente";
			// 
			// frmCred_ReporteVerificacionTelefonica
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1192, 366);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.dgVerificacion);
			this.Controls.Add(this.txtIdCre_Solicitud);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.txtNENombres);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNEIdentificacion);
			this.Name = "frmCred_ReporteVerificacionTelefonica";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Reporte Verificacion Telefonica";
			this.Load += new System.EventHandler(this.frmCred_ReporteVerificacionTelefonica_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNEIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNENombres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Solicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dgVerificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCred_ReporteVerificacionTelefonica_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083405");

			if (!Funcion.Permiso("985", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Verificacion Telefonica", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
		  bCarga = true;
			this.optEstados.Value = 0;
			
			if (bpantalla)
			{
				this.txtNEIdentificacion.Text= scedula;
				this.txtNEIdentificacion.Enabled = false;
				this.btnVer_Click ( sender, e);
				this.btnImprimir.Enabled = false;
				this.dgVerificacion.DisplayLayout.Bands[0].Columns["Sel"].Hidden = true;

			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void txtNEIdentificacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtNEIdentificacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
     	if(e.KeyCode == Keys.Enter) this.btnVer_Click(sender,e);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbTexto(this.txtNEIdentificacion, 3, 10, "Cedula del Cliente")) 
			{ 
				this.txtNEIdentificacion.Focus();
				return;
			}
			
			string sSQL = string.Format("Exec Cre_BuscaNombreVerificacionTelefonica '{0}', {1}", this.txtNEIdentificacion.Text.ToString(), (int) this.optEstados.Value);
			SqlDataReader drD= Funcion.rEscalarSQL(cdsSeteoF , sSQL, true);
			drD.Read();
			if (drD.HasRows)
			{
				if (drD.GetValue (0) != System.DBNull.Value)	this.txtIdCre_Solicitud.Value  = drD.GetInt32 (0);
				if (drD.GetValue (1) != System.DBNull.Value)this.txtNENombres.Text = drD.GetString (1);	
			}
			drD.Close();

			#region Referencias
			string sqlGRP = string.Format ("Exec Cre_ListaVerificacionTelefonica '{0}', {1}", 
				this.txtNEIdentificacion.Text.ToString(), (int) this.optEstados.Value);  
			this.dgVerificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sqlGRP);

			if (	this.dgVerificacion.Rows.Count > 0)this.dgVerificacion.Focus(); else this.txtNEIdentificacion.Focus();

			#endregion Referencias 
		}

		private void dgVerificacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.dgVerificacion);
		}

		private void dgVerificacion_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

		}

		private void dgVerificacion_CellChange_1(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
				string sSQL = "";
			
			if((bool)e.Cell.Row.Cells["Sel"].Value)
			{
				DateTime dtFecha = DateTime.Today.AddDays(-30);
			if(Convert.ToDateTime(e.Cell.Row.Cells["Fecha"].Value) > dtFecha)
			{
				MessageBox.Show("Esta verificacion esta caducada", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dgVerificacion.Rows[e.Cell.Row.Index].Cells["Sel"].Value = false;
								}
			}
			else{

//				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
//				{
//					oConexion.Open();
//
//					SqlTransaction oTransaction = null;
//					SqlCommand oCmdActualiza = oConexion.CreateCommand();
//					oCmdActualiza.Connection = oConexion;
//					oCmdActualiza.CommandType = CommandType.Text;
//					oCmdActualiza.CommandTimeout = 120;
//
//
//					if ((bool)e.Cell.Row.Cells["Sel"].Value)
//					{
//
//						try
//						{
//							oTransaction = oConexion.BeginTransaction();
//							oCmdActualiza.Transaction = oTransaction;
//
//							sSQL = string.Format("Exec Cre_ActualizaSeleccionVeriTelef {0}, {1}", (bool)e.Cell.Row.Cells["Sel"].Value, (int)e.Cell.Row.Cells["idCre_VerificacionTelefonica"].Value);
//							oCmdActualiza.CommandText = sSQL;
//							oCmdActualiza.ExecuteNonQuery();
//
//							oTransaction.Commit();
//							MessageBox.Show("Datos Actualizados", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
//						}
//						catch (Exception ex)
//						{
//							oTransaction.Rollback();
//
//							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);	
//							return;
//						}
//						finally
//						{
//							oConexion.Close();
//						}
//
////		
//					}
//				}
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{

			if (this.txtNEIdentificacion.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Numero de Cedula del Cliente", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
		if (this.dgVerificacion.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros de verificacion telefonica", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			return;
			}
		
			int iOrigen = 0;
			if(this.optEstados.CheckedIndex == 1)iOrigen = 2;

				ParameterFields paramFields = new ParameterFields();

				ParameterField pfCapital = new ParameterField();
				ParameterDiscreteValue ValCapital = new ParameterDiscreteValue();
				pfCapital.ParameterFieldName = "@Numero";
				ValCapital.Value = this.txtNEIdentificacion.Text.ToString();
				pfCapital.CurrentValues.Add(ValCapital);
				paramFields.Add(pfCapital);

				ParameterField pfFecha = new ParameterField();
				ParameterDiscreteValue ValFecha = new ParameterDiscreteValue();
				pfFecha.ParameterFieldName = "@idCre_Solicitud";
				ValFecha.Value = (int)this.txtIdCre_Solicitud.Value;
				pfFecha.CurrentValues.Add(ValFecha);
				paramFields.Add(pfFecha);

				ParameterField pfOrigen = new ParameterField();
				ParameterDiscreteValue ValOrigen = new ParameterDiscreteValue();
				pfOrigen.ParameterFieldName = "@Origen";
				ValOrigen.Value = iOrigen;
				pfOrigen.CurrentValues.Add(ValOrigen);
				paramFields.Add(pfOrigen);

				Reporte Reportes = new Reporte("RepVerificacionTelefonica.rpt");
			  Reportes.crVista.ParameterFieldInfo = paramFields;
				Reportes.MdiParent = this.MdiParent;
			
				Reportes.Show();
			
		}
	
		class validar 
		{
			public static void Solonumeros (KeyPressEventArgs Nu)
			{
				if(char.IsDigit(Nu.KeyChar))
				{
			
					Nu.Handled = false;
			
				}
				else if (char.IsControl(Nu.KeyChar))
				{
			
					Nu.Handled = false;			} 
				else
				{
					Nu.Handled= true;
				}
			}
		}

	}
}
