using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for DBUsuarioSeguridad.
	/// </summary>
	public class DBUsuarioSeguridad : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsUsuario;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraEditors.LabelControl labelControl2;
		private DevExpress.XtraEditors.LabelControl labelControl3;
		private DevExpress.XtraEditors.LabelControl labelControl4;
		private DevExpress.XtraEditors.LabelControl labelControl5;
		private DevExpress.XtraEditors.LabelControl labelControl6;
		private DevExpress.XtraEditors.LabelControl labelControl7;
		private DevExpress.XtraEditors.SpinEdit spnAncho;
		private DevExpress.XtraEditors.SpinEdit spnTexto;
		private DevExpress.XtraEditors.SpinEdit spnNumero;
		private DevExpress.XtraEditors.SpinEdit spnEspeciales;
		private DevExpress.XtraEditors.SpinEdit spnGrupo;
		private DevExpress.XtraEditors.SpinEdit spnDiasCaduca;
		private DevExpress.XtraEditors.CheckEdit chkPrimerIngreso;
		private System.Windows.Forms.TextBox txtId;
		private DevExpress.XtraEditors.LabelControl labelControl8;
		private DevExpress.XtraEditors.SpinEdit spnMayuscula;
		private System.ComponentModel.IContainer components;

		public DBUsuarioSeguridad()
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
			this.cdsUsuario = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.spnAncho = new DevExpress.XtraEditors.SpinEdit();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
			this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
			this.spnTexto = new DevExpress.XtraEditors.SpinEdit();
			this.spnNumero = new DevExpress.XtraEditors.SpinEdit();
			this.spnEspeciales = new DevExpress.XtraEditors.SpinEdit();
			this.spnGrupo = new DevExpress.XtraEditors.SpinEdit();
			this.labelControl6 = new DevExpress.XtraEditors.LabelControl();
			this.spnDiasCaduca = new DevExpress.XtraEditors.SpinEdit();
			this.labelControl7 = new DevExpress.XtraEditors.LabelControl();
			this.chkPrimerIngreso = new DevExpress.XtraEditors.CheckEdit();
			this.txtId = new System.Windows.Forms.TextBox();
			this.spnMayuscula = new DevExpress.XtraEditors.SpinEdit();
			this.labelControl8 = new DevExpress.XtraEditors.LabelControl();
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnAncho.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTexto.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumero.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnEspeciales.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDiasCaduca.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chkPrimerIngreso.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnMayuscula.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsUsuario
			// 
			this.cdsUsuario.BindingContextCtrl = this;
			this.cdsUsuario.DataLibrary = "LibDBEmpresa";
			this.cdsUsuario.DataLibraryUrl = "";
			this.cdsUsuario.DataSetDef = "dsEmpresa";
			this.cdsUsuario.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsUsuario.Name = "cdsUsuario";
			this.cdsUsuario.SchemaClassName = "LibDBEmpresa.DataClass";
			this.cdsUsuario.SchemaDef = null;
			this.cdsUsuario.BeforeFill += new C1.Data.FillEventHandler(this.cdsUsuario_BeforeFill);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "UsuarioSeguridad";
			this.barraDato1.DataNombreId = "idUsuarioSeguridad";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsUsuario;
			this.barraDato1.DataTabla = "UsuarioSeguridad";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(29, 421);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(326, 29);
			this.barraDato1.TabIndex = 7;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// spnAncho
			// 
			this.spnAncho.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsUsuario, "UsuarioSeguridad.Ancho"));
			this.spnAncho.EditValue = new System.Decimal(new int[] {
																															 0,
																															 0,
																															 0,
																															 0});
			this.spnAncho.Location = new System.Drawing.Point(307, 128);
			this.spnAncho.Name = "spnAncho";
			// 
			// spnAncho.Properties
			// 
			this.spnAncho.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnAncho.Properties.IsFloatValue = false;
			this.spnAncho.Properties.Mask.EditMask = "N00";
			this.spnAncho.Properties.MaxValue = new System.Decimal(new int[] {
																																				 99,
																																				 0,
																																				 0,
																																				 0});
			this.spnAncho.Size = new System.Drawing.Size(67, 22);
			this.spnAncho.TabIndex = 8;
			// 
			// labelControl1
			// 
			this.labelControl1.Location = new System.Drawing.Point(38, 128);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(183, 18);
			this.labelControl1.TabIndex = 12;
			this.labelControl1.Text = "Longitud mínima de Contraseña";
			// 
			// labelControl2
			// 
			this.labelControl2.Location = new System.Drawing.Point(38, 201);
			this.labelControl2.Name = "labelControl2";
			this.labelControl2.Size = new System.Drawing.Size(212, 18);
			this.labelControl2.TabIndex = 13;
			this.labelControl2.Text = "Cantidad de minúsculas obligatorias";
			// 
			// labelControl3
			// 
			this.labelControl3.Location = new System.Drawing.Point(38, 238);
			this.labelControl3.Name = "labelControl3";
			this.labelControl3.Size = new System.Drawing.Size(202, 18);
			this.labelControl3.TabIndex = 14;
			this.labelControl3.Text = "Cantidad de números obligatorios";
			// 
			// labelControl4
			// 
			this.labelControl4.Location = new System.Drawing.Point(38, 274);
			this.labelControl4.Name = "labelControl4";
			this.labelControl4.Size = new System.Drawing.Size(202, 19);
			this.labelControl4.TabIndex = 15;
			this.labelControl4.Text = "Cantidad de simbolos obligatorios";
			// 
			// labelControl5
			// 
			this.labelControl5.Appearance.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelControl5.Appearance.Options.UseFont = true;
			this.labelControl5.Location = new System.Drawing.Point(29, 18);
			this.labelControl5.Name = "labelControl5";
			this.labelControl5.Size = new System.Drawing.Size(365, 55);
			this.labelControl5.TabIndex = 16;
			this.labelControl5.Text = "Configuración de formación de claves para grupos de usuarios. Si no tiene un requ" +
				"isito específico déjelo en cero";
			// 
			// spnTexto
			// 
			this.spnTexto.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsUsuario, "UsuarioSeguridad.Minuscula"));
			this.spnTexto.EditValue = new System.Decimal(new int[] {
																															 0,
																															 0,
																															 0,
																															 0});
			this.spnTexto.Location = new System.Drawing.Point(307, 201);
			this.spnTexto.Name = "spnTexto";
			// 
			// spnTexto.Properties
			// 
			this.spnTexto.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnTexto.Properties.IsFloatValue = false;
			this.spnTexto.Properties.Mask.EditMask = "N00";
			this.spnTexto.Properties.MaxValue = new System.Decimal(new int[] {
																																				 99,
																																				 0,
																																				 0,
																																				 0});
			this.spnTexto.Size = new System.Drawing.Size(67, 22);
			this.spnTexto.TabIndex = 17;
			// 
			// spnNumero
			// 
			this.spnNumero.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsUsuario, "UsuarioSeguridad.Numero"));
			this.spnNumero.EditValue = new System.Decimal(new int[] {
																																0,
																																0,
																																0,
																																0});
			this.spnNumero.Location = new System.Drawing.Point(307, 238);
			this.spnNumero.Name = "spnNumero";
			// 
			// spnNumero.Properties
			// 
			this.spnNumero.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnNumero.Properties.IsFloatValue = false;
			this.spnNumero.Properties.Mask.EditMask = "N00";
			this.spnNumero.Properties.MaxValue = new System.Decimal(new int[] {
																																					99,
																																					0,
																																					0,
																																					0});
			this.spnNumero.Size = new System.Drawing.Size(67, 22);
			this.spnNumero.TabIndex = 18;
			// 
			// spnEspeciales
			// 
			this.spnEspeciales.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsUsuario, "UsuarioSeguridad.Especiales"));
			this.spnEspeciales.EditValue = new System.Decimal(new int[] {
																																		0,
																																		0,
																																		0,
																																		0});
			this.spnEspeciales.Location = new System.Drawing.Point(307, 274);
			this.spnEspeciales.Name = "spnEspeciales";
			// 
			// spnEspeciales.Properties
			// 
			this.spnEspeciales.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnEspeciales.Properties.IsFloatValue = false;
			this.spnEspeciales.Properties.Mask.EditMask = "N00";
			this.spnEspeciales.Properties.MaxValue = new System.Decimal(new int[] {
																																							99,
																																							0,
																																							0,
																																							0});
			this.spnEspeciales.Size = new System.Drawing.Size(67, 22);
			this.spnEspeciales.TabIndex = 19;
			this.spnEspeciales.ToolTip = "Simbolos válidos  [|@°!\"#$%&/()=?¡<>¿+*{}_-,.;:]";
			// 
			// spnGrupo
			// 
			this.spnGrupo.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsUsuario, "UsuarioSeguridad.idGrupo"));
			this.spnGrupo.EditValue = new System.Decimal(new int[] {
																															 1,
																															 0,
																															 0,
																															 0});
			this.spnGrupo.Location = new System.Drawing.Point(307, 91);
			this.spnGrupo.Name = "spnGrupo";
			// 
			// spnGrupo.Properties
			// 
			this.spnGrupo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnGrupo.Properties.IsFloatValue = false;
			this.spnGrupo.Properties.Mask.EditMask = "N00";
			this.spnGrupo.Properties.MaxValue = new System.Decimal(new int[] {
																																				 20,
																																				 0,
																																				 0,
																																				 0});
			this.spnGrupo.Properties.MinValue = new System.Decimal(new int[] {
																																				 1,
																																				 0,
																																				 0,
																																				 0});
			this.spnGrupo.Size = new System.Drawing.Size(67, 22);
			this.spnGrupo.TabIndex = 20;
			// 
			// labelControl6
			// 
			this.labelControl6.Location = new System.Drawing.Point(38, 91);
			this.labelControl6.Name = "labelControl6";
			this.labelControl6.Size = new System.Drawing.Size(192, 16);
			this.labelControl6.TabIndex = 21;
			this.labelControl6.Text = "Grupo al que aplica condiciones";
			// 
			// spnDiasCaduca
			// 
			this.spnDiasCaduca.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsUsuario, "UsuarioSeguridad.DiasCaduca"));
			this.spnDiasCaduca.EditValue = new System.Decimal(new int[] {
																																		0,
																																		0,
																																		0,
																																		0});
			this.spnDiasCaduca.Location = new System.Drawing.Point(307, 311);
			this.spnDiasCaduca.Name = "spnDiasCaduca";
			// 
			// spnDiasCaduca.Properties
			// 
			this.spnDiasCaduca.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnDiasCaduca.Properties.IsFloatValue = false;
			this.spnDiasCaduca.Properties.Mask.EditMask = "N00";
			this.spnDiasCaduca.Properties.MaxValue = new System.Decimal(new int[] {
																																							999,
																																							0,
																																							0,
																																							0});
			this.spnDiasCaduca.Size = new System.Drawing.Size(67, 22);
			this.spnDiasCaduca.TabIndex = 23;
			// 
			// labelControl7
			// 
			this.labelControl7.Location = new System.Drawing.Point(38, 311);
			this.labelControl7.Name = "labelControl7";
			this.labelControl7.Size = new System.Drawing.Size(212, 18);
			this.labelControl7.TabIndex = 22;
			this.labelControl7.Text = "Días a los que caduca la contraseña";
			// 
			// chkPrimerIngreso
			// 
			this.chkPrimerIngreso.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsUsuario, "UsuarioSeguridad.ClaveInicio"));
			this.chkPrimerIngreso.Location = new System.Drawing.Point(38, 366);
			this.chkPrimerIngreso.Name = "chkPrimerIngreso";
			// 
			// chkPrimerIngreso.Properties
			// 
			this.chkPrimerIngreso.Properties.Caption = "Ingreso de contraseña el primer ingreso";
			this.chkPrimerIngreso.Size = new System.Drawing.Size(336, 21);
			this.chkPrimerIngreso.TabIndex = 24;
			// 
			// txtId
			// 
			this.txtId.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsUsuario, "UsuarioSeguridad.idUsuarioSeguridad"));
			this.txtId.Location = new System.Drawing.Point(365, 421);
			this.txtId.Name = "txtId";
			this.txtId.Size = new System.Drawing.Size(57, 22);
			this.txtId.TabIndex = 25;
			this.txtId.Text = "Id";
			// 
			// spnMayuscula
			// 
			this.spnMayuscula.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsUsuario, "UsuarioSeguridad.Texto"));
			this.spnMayuscula.EditValue = new System.Decimal(new int[] {
																																	 0,
																																	 0,
																																	 0,
																																	 0});
			this.spnMayuscula.Location = new System.Drawing.Point(307, 165);
			this.spnMayuscula.Name = "spnMayuscula";
			// 
			// spnMayuscula.Properties
			// 
			this.spnMayuscula.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				 new DevExpress.XtraEditors.Controls.EditorButton()});
			this.spnMayuscula.Properties.IsFloatValue = false;
			this.spnMayuscula.Properties.Mask.EditMask = "N00";
			this.spnMayuscula.Properties.MaxValue = new System.Decimal(new int[] {
																																						 99,
																																						 0,
																																						 0,
																																						 0});
			this.spnMayuscula.Size = new System.Drawing.Size(67, 22);
			this.spnMayuscula.TabIndex = 27;
			// 
			// labelControl8
			// 
			this.labelControl8.Location = new System.Drawing.Point(38, 165);
			this.labelControl8.Name = "labelControl8";
			this.labelControl8.Size = new System.Drawing.Size(231, 18);
			this.labelControl8.TabIndex = 26;
			this.labelControl8.Text = "Cantidad de Mayusculas obligatorias";
			// 
			// DBUsuarioSeguridad
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 16);
			this.ClientSize = new System.Drawing.Size(422, 464);
			this.Controls.Add(this.spnMayuscula);
			this.Controls.Add(this.labelControl8);
			this.Controls.Add(this.txtId);
			this.Controls.Add(this.chkPrimerIngreso);
			this.Controls.Add(this.spnDiasCaduca);
			this.Controls.Add(this.labelControl7);
			this.Controls.Add(this.labelControl6);
			this.Controls.Add(this.spnGrupo);
			this.Controls.Add(this.spnEspeciales);
			this.Controls.Add(this.spnNumero);
			this.Controls.Add(this.spnTexto);
			this.Controls.Add(this.labelControl5);
			this.Controls.Add(this.labelControl4);
			this.Controls.Add(this.labelControl3);
			this.Controls.Add(this.labelControl2);
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.spnAncho);
			this.Controls.Add(this.barraDato1);
			this.Name = "DBUsuarioSeguridad";
			this.Text = "Especificaciones de Seguridad";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.DBUsuarioSeguridad_Closing);
			this.Load += new System.EventHandler(this.DBUsuarioSeguridad_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnAncho.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTexto.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumero.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnEspeciales.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDiasCaduca.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chkPrimerIngreso.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnMayuscula.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			
		}

		private bool Verifica()
		{
			int iCar = int.Parse(spnTexto.EditValue.ToString());
			int iNum = int.Parse(spnNumero.EditValue.ToString());
			int iEsp = int.Parse(spnEspeciales.EditValue.ToString());
			int iMayuscula = int.Parse(spnMayuscula.EditValue.ToString());
			int iAncho = int.Parse(spnAncho.EditValue.ToString());

			int iCalculo = iCar + iNum + iEsp + iMayuscula;
			if (iCalculo > iAncho)
			{
				string stMensaje = "Ancho debe ser mayor a: " + iCalculo;
				MessageBox.Show(stMensaje, "Informacion");
				spnAncho.SelectAll();
				return false;
			}
			int idGrupo = int.Parse(spnGrupo.EditValue.ToString());
			int idUsuarioSeguridad = int.Parse(txtId.Text);
			string stSelect = string.Format(
				"Select Count(*) From UsuarioSeguridad Where idGrupo = {0} And idUsuarioSeguridad != {1}",
				idGrupo, idUsuarioSeguridad);
			int iCuenta = Funcion.iEscalarSQL(cdsUsuario, stSelect);
			if (iCuenta > 0)
			{
				MessageBox.Show("Grupo ya existe", "Informacion");
				spnGrupo.SelectAll();
				return false;
			}

			return true;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			barraDato1.GrabaRegistro();
		}

//		Acceso miAcceso;
		private void DBUsuarioSeguridad_Load(object sender, System.EventArgs e)
		{
			#region Auditoria
			string stAudita = "Exec AuditaCrear 53, 6, 'Especif_Seguridad'";
			Funcion.EjecutaSQL(cdsUsuario, stAudita, true);
			#endregion Auditoria
			#region Seguridad
//			miAcceso =new Acceso(cdsUsuario,  "050A21");
//			if (!miAcceso.Eliminar) barraDato1.sbtBorrar.Enabled = false;
//			if (!miAcceso.Editar) barraDato1.sbtEditar.Enabled = false;
//			if (!miAcceso.Nuevo) barraDato1.sbtNuevo.Enabled = false;
			#endregion Seguridad

			txtId.Width = 0;
			barraDato1.HabilitaControles(false);
			barraDato1.ProximoId(4);
		}

		private void cdsUsuario_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsUsuario.Schema.Connections[0].ConnectionString = miFuncion.AccesoEmpresaBase();
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("¿Desea Borrar el Registro?\n" +
				"Los usuarios de este grupo no tendrán restricciones de claves.", "Confirmación", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2) ) return;
			barraDato1.BorraRegistro();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			barraDato1.CrearRegistro();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			barraDato1.EditarRegistro();
		}

		private void DBUsuarioSeguridad_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Especif_Seguridad'";
			Funcion.EjecutaSQL(cdsUsuario, stAudita, true);
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
			DBUsuario miUsuario = new DBUsuario();
			miUsuario.MdiParent = this.MdiParent;
			miUsuario.Show();
		}
	}
}

