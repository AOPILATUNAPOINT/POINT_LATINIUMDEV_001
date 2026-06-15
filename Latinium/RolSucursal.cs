using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolSucursal.
	/// </summary>
	public class RolSucursal : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsSucursal;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSucursal;
		private C1.Data.C1DataView cdvTipoEmpleador;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoEmp;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtOrigenFondos;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigoSeguro;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroSeguro;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private System.Windows.Forms.Label label6;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolSucursal()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolTipoEmpleador", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipoEmpleador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			this.cdsSucursal = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.txtSucursal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdvTipoEmpleador = new C1.Data.C1DataView();
			this.cmbTipoEmp = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.txtOrigenFondos = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtCodigoSeguro = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtNumeroSeguro = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			((System.ComponentModel.ISupportInitialize)(this.cdsSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoEmpleador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoEmp)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigenFondos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoSeguro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroSeguro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsSucursal
			// 
			this.cdsSucursal.BindingContextCtrl = this;
			this.cdsSucursal.DataLibrary = "LibPersonal";
			this.cdsSucursal.DataLibraryUrl = "";
			this.cdsSucursal.DataSetDef = "dsRolTablas";
			this.cdsSucursal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSucursal.Name = "cdsSucursal";
			this.cdsSucursal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsSucursal.SchemaDef = null;
			this.cdsSucursal.BeforeFill += new C1.Data.FillEventHandler(this.cdsSucursal_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Sucursal:";
			// 
			// txtSucursal
			// 
			this.txtSucursal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSucursal, "RolSucursal.Sucursal"));
			this.txtSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSucursal.Location = new System.Drawing.Point(80, 16);
			this.txtSucursal.Name = "txtSucursal";
			this.txtSucursal.Size = new System.Drawing.Size(240, 21);
			this.txtSucursal.TabIndex = 1;
			// 
			// cdvTipoEmpleador
			// 
			this.cdvTipoEmpleador.BindingContextCtrl = this;
			this.cdvTipoEmpleador.DataSet = this.cdsSucursal;
			this.cdvTipoEmpleador.TableName = "";
			this.cdvTipoEmpleador.TableViewName = "RolTipoEmpleador";
			// 
			// cmbTipoEmp
			// 
			this.cmbTipoEmp.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoEmp.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoEmp.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSucursal, "RolSucursal.idRolTipoEmpleador"));
			this.cmbTipoEmp.DataSource = this.cdvTipoEmpleador;
			this.cmbTipoEmp.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 56;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 118;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbTipoEmp.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoEmp.DisplayMember = "Tipo";
			this.cmbTipoEmp.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoEmp.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoEmp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoEmp.Location = new System.Drawing.Point(128, 48);
			this.cmbTipoEmp.Name = "cmbTipoEmp";
			this.cmbTipoEmp.Size = new System.Drawing.Size(192, 21);
			this.cmbTipoEmp.TabIndex = 3;
			this.cmbTipoEmp.ValueMember = "idRolTipoEmpleador";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 48);
			this.label2.Name = "label2";
			this.label2.TabIndex = 4;
			this.label2.Text = "Tipo de Empleador";
			// 
			// txtOrigenFondos
			// 
			this.txtOrigenFondos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSucursal, "RolSucursal.OrigenFondos"));
			this.txtOrigenFondos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOrigenFondos.Location = new System.Drawing.Point(128, 80);
			this.txtOrigenFondos.Name = "txtOrigenFondos";
			this.txtOrigenFondos.Size = new System.Drawing.Size(48, 21);
			this.txtOrigenFondos.TabIndex = 6;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 80);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(98, 16);
			this.label3.TabIndex = 5;
			this.label3.Text = "Origen de Fondos:";
			// 
			// txtCodigoSeguro
			// 
			this.txtCodigoSeguro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSucursal, "RolSucursal.CodigoSeguro"));
			this.txtCodigoSeguro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigoSeguro.Location = new System.Drawing.Point(128, 112);
			this.txtCodigoSeguro.Name = "txtCodigoSeguro";
			this.txtCodigoSeguro.Size = new System.Drawing.Size(48, 21);
			this.txtCodigoSeguro.TabIndex = 8;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 112);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(95, 16);
			this.label4.TabIndex = 7;
			this.label4.Text = "Código de Seguro";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(16, 144);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(103, 16);
			this.label5.TabIndex = 10;
			this.label5.Text = "Número de Seguro:";
			// 
			// txtNumeroSeguro
			// 
			this.txtNumeroSeguro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSucursal, "RolSucursal.Numero"));
			this.txtNumeroSeguro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroSeguro.Location = new System.Drawing.Point(128, 144);
			this.txtNumeroSeguro.Name = "txtNumeroSeguro";
			this.txtNumeroSeguro.Size = new System.Drawing.Size(48, 21);
			this.txtNumeroSeguro.TabIndex = 11;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSucursal, "RolSucursal.Codigo"));
			this.ultraTextEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor1.Location = new System.Drawing.Point(128, 176);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(48, 21);
			this.ultraTextEditor1.TabIndex = 13;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 176);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(106, 16);
			this.label6.TabIndex = 12;
			this.label6.Text = "Codigo de Sucursal:";
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "RolSucursal";
			this.barraDato1.DataNombreId = "idRolSucursal";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsSucursal;
			this.barraDato1.DataTabla = "RolSucursal";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(16, 216);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 16;
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
			// RolSucursal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(344, 254);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNumeroSeguro);
			this.Controls.Add(this.txtCodigoSeguro);
			this.Controls.Add(this.txtOrigenFondos);
			this.Controls.Add(this.cmbTipoEmp);
			this.Controls.Add(this.txtSucursal);
			this.MaximizeBox = false;
			this.Name = "RolSucursal";
			this.Text = "Sucursal";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RolSucursal_Closing);
			this.Load += new System.EventHandler(this.RolSucursal_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipoEmpleador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoEmp)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigenFondos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigoSeguro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroSeguro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void RolSucursal_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Rol Sucursal'";
			Funcion.EjecutaSQL(cdsSucursal, stAudita, true);
			Cursor = Cursors.Default;
			this.barraDato1.HabilitaControles(false);
			this.barraDato1.PosUltima();
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea borrar ésta sucursal?",
				"Confirmación", MessageBoxButtons.YesNo,
				MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
				== DialogResult.No) return;
			this.barraDato1.BorraRegistro();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.BorraErrores();
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
		}

		private void BorraErrores()
		{
			this.errorProvider.SetError(this.txtSucursal, "");
			this.errorProvider.SetError(this.txtCodigoSeguro, "");
			this.errorProvider.SetError(this.txtOrigenFondos, "");
			this.errorProvider.SetError(this.txtNumeroSeguro, "");
			this.errorProvider.SetError(this.cmbTipoEmp, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			this.BorraErrores();
			if (this.txtSucursal.Value == null || this.txtSucursal.Value.ToString().Trim().Length == 0)
			{
				bOk = false;
				this.errorProvider.SetError(this.txtSucursal, "Ingrese Nombre de Sucursal");
			}
			if (this.txtCodigoSeguro.Value == null || this.txtCodigoSeguro.Value.ToString().Trim().Length == 0)
			{
				bOk = false;
				this.errorProvider.SetError(this.txtCodigoSeguro, "Ingreso Código del Seguro.\n"
					+ "R Seguro Social Vigente. M Seguro Mixto");
			}
			if (this.txtNumeroSeguro.Value == null || this.txtNumeroSeguro.Value.ToString().Trim().Length == 0)
			{
				bOk = false;
				this.errorProvider.SetError(this.txtNumeroSeguro, "Ingrese Numero de Sucursal");
			}
			if (this.txtOrigenFondos.Value == null || this.txtOrigenFondos.Value.ToString().Trim().Length == 0)
			{
				bOk = false;
				this.errorProvider.SetError(this.txtOrigenFondos, "Ingrese Origen de Fondos.\n"+
					"P Fondos Propios. E Presupuesto del Estado");
			}
			if (this.cmbTipoEmp.Text.ToString().Trim().Length == 0)
			{
				bOk = false;
				this.errorProvider.SetError(this.cmbTipoEmp, "Ingrese Tipo de Empleador.");
			}
			return bOk;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			this.barraDato1.GrabaRegistro();
		}

		private void cdsSucursal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSucursal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void RolSucursal_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Rol Sucursal'";
			Funcion.EjecutaSQL(cdsSucursal, stAudita, true);
		}
	}
}
