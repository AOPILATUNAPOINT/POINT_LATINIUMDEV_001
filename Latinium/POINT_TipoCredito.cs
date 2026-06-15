using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de POINT_TipoCredito.
	/// </summary>
	public class POINT_TipoCredito : System.Windows.Forms.Form
	{
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private C1.Data.C1DataSet cdsContacto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtInteres;
		private DevExpress.XtraEditors.RadioGroup rbgTipoCredito;
		private DevExpress.XtraEditors.RadioGroup rbgAfectaa;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public POINT_TipoCredito()
		{
			InitializeComponent();

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
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsContacto = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDiasCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtNumCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtInteres = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.rbgTipoCredito = new DevExpress.XtraEditors.RadioGroup();
			this.rbgAfectaa = new DevExpress.XtraEditors.RadioGroup();
			((System.ComponentModel.ISupportInitialize)(this.cdsContacto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbgTipoCredito.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbgAfectaa.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "TIPOCRED";
			this.barraDato1.DataNombreId = "idTipoCredito";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsContacto;
			this.barraDato1.DataTabla = "TIPOCRED";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(24, 272);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 78;
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cdsContacto
			// 
			this.cdsContacto.BindingContextCtrl = this;
			this.cdsContacto.DataLibrary = "LibFacturacion";
			this.cdsContacto.DataLibraryUrl = "";
			this.cdsContacto.DataSetDef = "dsClienteTabla";
			this.cdsContacto.FillOnRequest = false;
			this.cdsContacto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsContacto.Name = "cdsContacto";
			this.cdsContacto.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsContacto.SchemaDef = null;
			this.cdsContacto.BeforeFill += new C1.Data.FillEventHandler(this.cdsContacto_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(48, 18);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 80;
			this.label1.Text = "Código:";
			// 
			// txtCodigo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance1;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContacto, "TIPOCRED.Codigo"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(96, 16);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(40, 21);
			this.txtCodigo.TabIndex = 79;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(43, 50);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 16);
			this.label2.TabIndex = 82;
			this.label2.Text = "Nombre:";
			// 
			// txtNombre
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance2;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContacto, "TIPOCRED.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(96, 48);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(288, 21);
			this.txtNombre.TabIndex = 81;
			// 
			// txtDiasCredito
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasCredito.Appearance = appearance3;
			this.txtDiasCredito.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContacto, "TIPOCRED.Dias"));
			this.txtDiasCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasCredito.Location = new System.Drawing.Point(96, 80);
			this.txtDiasCredito.MinValue = 0;
			this.txtDiasCredito.Name = "txtDiasCredito";
			this.txtDiasCredito.PromptChar = ' ';
			this.txtDiasCredito.Size = new System.Drawing.Size(72, 21);
			this.txtDiasCredito.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDiasCredito.TabIndex = 109;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(21, 82);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(70, 16);
			this.label12.TabIndex = 110;
			this.label12.Text = "Días Crédito:";
			// 
			// txtNumCuotas
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumCuotas.Appearance = appearance4;
			this.txtNumCuotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContacto, "TIPOCRED.Giros"));
			this.txtNumCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCuotas.Location = new System.Drawing.Point(96, 112);
			this.txtNumCuotas.MinValue = 0;
			this.txtNumCuotas.Name = "txtNumCuotas";
			this.txtNumCuotas.PromptChar = ' ';
			this.txtNumCuotas.Size = new System.Drawing.Size(72, 21);
			this.txtNumCuotas.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNumCuotas.TabIndex = 111;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(38, 114);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(53, 16);
			this.label3.TabIndex = 112;
			this.label3.Text = "# Cuotas:";
			// 
			// txtInteres
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInteres.Appearance = appearance5;
			this.txtInteres.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContacto, "TIPOCRED.PctIneteres"));
			this.txtInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInteres.Location = new System.Drawing.Point(96, 144);
			this.txtInteres.MinValue = 0;
			this.txtInteres.Name = "txtInteres";
			this.txtInteres.PromptChar = ' ';
			this.txtInteres.Size = new System.Drawing.Size(72, 21);
			this.txtInteres.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtInteres.TabIndex = 113;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(36, 146);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(55, 16);
			this.label4.TabIndex = 114;
			this.label4.Text = "% Interes:";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(32, 184);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(66, 16);
			this.label5.TabIndex = 148;
			this.label5.Text = "Tipo Crédito";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(200, 184);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(50, 16);
			this.label6.TabIndex = 149;
			this.label6.Text = "Afecta A:";
			// 
			// rbgTipoCredito
			// 
			this.rbgTipoCredito.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsContacto, "TIPOCRED.Tipo"));
			this.rbgTipoCredito.Location = new System.Drawing.Point(32, 200);
			this.rbgTipoCredito.Name = "rbgTipoCredito";
			// 
			// rbgTipoCredito.Properties
			// 
			this.rbgTipoCredito.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																					 new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Contado"),
																																																					 new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Crédto")});
			this.rbgTipoCredito.Size = new System.Drawing.Size(144, 32);
			this.rbgTipoCredito.TabIndex = 150;
			// 
			// rbgAfectaa
			// 
			this.rbgAfectaa.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsContacto, "TIPOCRED.AfectaA"));
			this.rbgAfectaa.Location = new System.Drawing.Point(200, 200);
			this.rbgAfectaa.Name = "rbgAfectaa";
			// 
			// rbgAfectaa.Properties
			// 
			this.rbgAfectaa.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			 new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Cliente"),
																																																			 new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Proveedor"),
																																																			 new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Ambos")});
			this.rbgAfectaa.Size = new System.Drawing.Size(208, 32);
			this.rbgAfectaa.TabIndex = 151;
			// 
			// POINT_TipoCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(440, 310);
			this.Controls.Add(this.rbgAfectaa);
			this.Controls.Add(this.rbgTipoCredito);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtInteres);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtNumCuotas);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtDiasCredito);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.barraDato1);
			this.Name = "POINT_TipoCredito";
			this.Text = "POINT_TipoCredito";
			this.Load += new System.EventHandler(this.POINT_TipoCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsContacto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbgTipoCredito.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbgAfectaa.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private C1.Data.C1DataRow drTipoCredto;

		private void cdsContacto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsContacto.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void POINT_TipoCredito_Load(object sender, System.EventArgs e)
		{
			this.barraDato1.HabilitaControles(false);
			this.barraDato1.PosUltima();
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			//drTipoCredto = cdsContacto.TableViews["TIPOCRE"].Rows[0];

		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			this.barraDato1.BorraRegistro();

		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			this.barraDato1.ProximoId(5);
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();

		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();

		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			this.barraDato1.GrabaRegistro();

		}
	}
}
