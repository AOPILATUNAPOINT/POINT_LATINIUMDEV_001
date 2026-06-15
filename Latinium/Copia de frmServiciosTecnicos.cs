using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmServiciosTecnicos.
	/// </summary>
	public class frmServiciosTecnicos : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdServicio;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsServicio;

		private C1.Data.C1DataRow drServicio;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsContactos;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmServiciosTecnicos()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ContactosServicioTecnico", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idContactoST");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idServicioTecnico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.txtIdServicio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsServicio = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtObservacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtTiempo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.label7 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsContactos = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtIdServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsServicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContactos)).BeginInit();
			this.SuspendLayout();
			// 
			// txtIdServicio
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			appearance1.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance1.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdServicio.Appearance = appearance1;
			this.txtIdServicio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsServicio, "ServicioTecnicoAut.idServicioTecnico"));
			this.txtIdServicio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdServicio.FormatString = "";
			this.txtIdServicio.Location = new System.Drawing.Point(438, 6);
			this.txtIdServicio.Name = "txtIdServicio";
			this.txtIdServicio.PromptChar = ' ';
			this.txtIdServicio.Size = new System.Drawing.Size(80, 21);
			this.txtIdServicio.TabIndex = 244;
			this.txtIdServicio.Visible = false;
			// 
			// cdsServicio
			// 
			this.cdsServicio.BindingContextCtrl = this;
			this.cdsServicio.DataLibrary = "LibPersonal";
			this.cdsServicio.DataLibraryUrl = "";
			this.cdsServicio.DataSetDef = "dsServicioTecnico";
			this.cdsServicio.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsServicio.Name = "cdsServicio";
			this.cdsServicio.SchemaClassName = "LibPersonal.DataClass";
			this.cdsServicio.SchemaDef = null;
			this.cdsServicio.BeforeFill += new C1.Data.FillEventHandler(this.cdsServicio_BeforeFill);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(10, 38);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(50, 16);
			this.label2.TabIndex = 241;
			this.label2.Text = "Dirección";
			// 
			// txtDireccion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance2;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsServicio, "ServicioTecnicoAut.Direcciòn"));
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Location = new System.Drawing.Point(126, 38);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(392, 22);
			this.txtDireccion.TabIndex = 240;
			this.txtDireccion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDireccion_KeyPress);
			this.txtDireccion.Validating += new System.ComponentModel.CancelEventHandler(this.txtDireccion_Validating);
			this.txtDireccion.ValueChanged += new System.EventHandler(this.txtDireccion_ValueChanged);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(10, 6);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(44, 16);
			this.label5.TabIndex = 237;
			this.label5.Text = "Nombre";
			// 
			// txtNombre
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance3;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsServicio, "ServicioTecnicoAut.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(126, 6);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(296, 22);
			this.txtNombre.TabIndex = 236;
			this.txtNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombre_KeyPress);
			this.txtNombre.Validating += new System.ComponentModel.CancelEventHandler(this.txtNombre_Validating);
			this.txtNombre.ValueChanged += new System.EventHandler(this.txtNombre_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(10, 104);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(76, 16);
			this.label3.TabIndex = 246;
			this.label3.Text = "Observaciones";
			// 
			// txtObservacion
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservacion.Appearance = appearance4;
			this.txtObservacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsServicio, "ServicioTecnicoAut.Observacion"));
			this.txtObservacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservacion.Location = new System.Drawing.Point(126, 104);
			this.txtObservacion.Multiline = true;
			this.txtObservacion.Name = "txtObservacion";
			this.txtObservacion.Size = new System.Drawing.Size(392, 64);
			this.txtObservacion.TabIndex = 245;
			this.txtObservacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtObservacion_KeyPress);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(10, 72);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(109, 16);
			this.label4.TabIndex = 248;
			this.label4.Text = "Tiempo de respuesta";
			// 
			// txtTiempo
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempo.Appearance = appearance5;
			this.txtTiempo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsServicio, "ServicioTecnicoAut.TiempoRespuesta"));
			this.txtTiempo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempo.FormatString = "";
			this.txtTiempo.Location = new System.Drawing.Point(126, 72);
			this.txtTiempo.Name = "txtTiempo";
			this.txtTiempo.PromptChar = ' ';
			this.txtTiempo.Size = new System.Drawing.Size(112, 22);
			this.txtTiempo.TabIndex = 249;
			this.txtTiempo.Click += new System.EventHandler(this.txtTiempo_Click);
			this.txtTiempo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTiempo_KeyPress);
			this.txtTiempo.Enter += new System.EventHandler(this.txtTiempo_Enter);
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.CausesValidation = false;
			this.barraDato1.DataMember = "ServicioTecnicoAut";
			this.barraDato1.DataNombreId = "idServicioTecnico";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsServicio;
			this.barraDato1.DataTabla = "ServicioTecnicoAut";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(8, 432);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(256, 26);
			this.barraDato1.TabIndex = 250;
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(254, 72);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(33, 16);
			this.label7.TabIndex = 251;
			this.label7.Text = "Horas";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "ContactosServicioTecnico";
			this.ultraGrid1.DataSource = this.cdsContactos;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 180;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 110;
			ultraGridColumn5.FieldLen = 9;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.FieldLen = 10;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			ultraGridBand1.Header.Caption = "Contactos";
			ultraGridBand1.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 10F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance9;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 176);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(512, 246);
			this.ultraGrid1.TabIndex = 252;
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cdsContactos
			// 
			this.cdsContactos.BindingContextCtrl = this;
			this.cdsContactos.DataLibrary = "LibPersonal";
			this.cdsContactos.DataLibraryUrl = "";
			this.cdsContactos.DataSetDef = "dsContactosServTec";
			this.cdsContactos.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsContactos.Name = "cdsContactos";
			this.cdsContactos.SchemaClassName = "LibPersonal.DataClass";
			this.cdsContactos.SchemaDef = null;
			this.cdsContactos.BeforeFill += new C1.Data.FillEventHandler(this.cdsContactos_BeforeFill);
			// 
			// frmServiciosTecnicos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(528, 464);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtTiempo);
			this.Controls.Add(this.txtObservacion);
			this.Controls.Add(this.txtIdServicio);
			this.Controls.Add(this.txtDireccion);
			this.Controls.Add(this.txtNombre);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmServiciosTecnicos";
			this.Text = "Servicio Técnico Autorizado";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmServiciosTecnicos_KeyDown);
			this.Load += new System.EventHandler(this.frmServiciosTecnicos_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtIdServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsServicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContactos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsServicio_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsServicio.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
			
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();

			this.txtNombre.Select();
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drServicio = cdsServicio.TableViews["ServicioTecnicoAut"].Rows[0];

				C1.Data.FilterCondition fcContactos = new C1.Data.FilterCondition(cdsContactos, "ContactosServicioTecnico", "idServicioTecnico = " + drServicio["idServicioTecnico"].ToString());
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcContactos);
				cdsContactos.Fill(fcTotal, false);
			}
			catch{};
		}

		private int WM_CLOSE = 0x0010;
		protected override void WndProc(ref System.Windows.Forms.Message msg)
		{
			if (msg.Msg == WM_CLOSE) 
			{
				ActiveControl.CausesValidation = false;
			}			
			base.WndProc(ref msg);			
		}

		private void frmServiciosTecnicos_Load(object sender, System.EventArgs e)
		{
			this.barraDato1.HabilitaControles(false);

			this.barraDato1.PosUltima();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.Select();
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
			this.txtNombre.Select();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.txtNombre.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba el nombre", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNombre.Select();
				return;
			}

			string sSQL = string.Format("Select Count(*) From ServicioTecnicoAut Where Nombre = '{0}' And idServicioTecnico <> {1}", this.txtNombre.Text.ToString(), (int)drServicio["idServicioTecnico"]);
			int iCuenta = Funcion.iEscalarSQL(cdsServicio, sSQL);
			if (iCuenta > 0)
			{
				MessageBox.Show(this.txtNombre.Text.ToString() + " Ya existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNombre.Select();
				return;
			}
		
			if (this.txtDireccion.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba la dirección.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDireccion.Select();
				return;
			}
		
			#region Contactos
			int iContContactos = this.ultraGrid1.Rows.Count;
			if (iContContactos < 1) 
			{
				MessageBox.Show("Debe ingresar un contacto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["Nombre"].Value == System.DBNull.Value || dr.Cells["Nombre"].Value.ToString() == "") 
				{
					MessageBox.Show("Ingrese un nombre de contacto en la Ficha Contactos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}

				if ((dr.Cells["Telefono"].Value == System.DBNull.Value || dr.Cells["Telefono"].Value.ToString() == "") && (dr.Cells["Celular"].Value == System.DBNull.Value || dr.Cells["Celular"].Value.ToString() == ""))
				{
					MessageBox.Show("Ingrese un número de contacto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}			
			}
			#endregion Contactos
			#endregion Validacion

			this.barraDato1.GrabaRegistro();

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				dr.Cells["idServicioTecnico"].Value = (int)drServicio["idServicioTecnico"];
			}

			cdsContactos.Update();
		}

		private void txtNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDireccion.Select();
		}

		private void txtDireccion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtTiempo.Select();
		}

		private void txtTiempo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtObservacion.Select();
		}

		private void txtObservacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.barraDato1.Focus();
		}

		private void txtNombre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!this.barraDato1.bEditar) return;
			if (this.txtNombre.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba el nombre", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNombre.Select();
				e.Cancel = true;
				return;
			}

			string sSQL = string.Format("Select Count(*) From ServicioTecnicoAut Where Nombre = '{0}' And idServicioTecnico <> {1}", this.txtNombre.Text.ToString(), (int)drServicio["idServicioTecnico"]);
      int iCuenta = Funcion.iEscalarSQL(cdsServicio, sSQL);
			if (iCuenta > 0)
			{
				MessageBox.Show(this.txtNombre.Text.ToString() + " Ya existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNombre.Select();
				e.Cancel = true;
				return;
			}
		}

		private void txtDireccion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!this.barraDato1.bEditar) return;

			if (this.txtDireccion.Text.ToString().Length == 0)
			{
				MessageBox.Show("Escriba la dirección.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDireccion.Select();
				e.Cancel = true;
				return;
			}
		}

		private void txtNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtTelefono_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using(frmBuscaServTecnico miBusqueda = new frmBuscaServTecnico())
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{					
					barraDato1.IdRegistro = miBusqueda.idServicio;
					this.barraDato1.ProximoId(5);
				}
			}
		}

		private void frmServiciosTecnicos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) 
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Deshacer(sender, e);

				}
				else
					Close();
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
					barraDato1.IdRegistro = (int)drServicio["idServicioTecnico"];
					barraDato1.ProximoId(5);					
				}
			}
		}

		private void txtDireccion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCelular_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtTiempo_Enter(object sender, System.EventArgs e)
		{
			this.txtTiempo.SelectAll();
		}

		private void cdsContactos_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsServicio.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			switch(e.KeyValue)
			{
					#region Flechas
				case (int) Keys.Up:
					if (ultraGrid1.ActiveCell.DroppedDown) return;

					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Left:
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (ultraGrid1.ActiveCell.DroppedDown) return;
			
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
					#endregion Flechas
					#region Enter
				case (int) Keys.Enter:
					ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
					ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
					if(ultraGrid1.ActiveCell.DroppedDown == false)
						ultraGrid1.ActiveCell.DroppedDown = true;
					//						ultraGrid3.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
					break;
					#endregion Enter
			}
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "TELEFONO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CELULAR")
				Funcion.CamposNumericos(sender, e);
		}

		private void txtTiempo_Click(object sender, System.EventArgs e)
		{
			this.txtTiempo.SelectAll();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}

