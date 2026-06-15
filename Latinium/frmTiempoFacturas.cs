using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTiempoFacturas.
	/// </summary>
	public class frmTiempoFacturas : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempo;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaSalida;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIngreso;
		private System.Windows.Forms.Button btnEditar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkActivo;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTiempoFacturas()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTiempoFacturas));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtTiempo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFechaSalida = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaIngreso = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnEditar = new System.Windows.Forms.Button();
			this.chkActivo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			this.SuspendLayout();
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
			// txtTiempo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempo.Appearance = appearance1;
			this.txtTiempo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempo.Enabled = false;
			this.txtTiempo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTiempo.FormatString = "#0";
			this.txtTiempo.Location = new System.Drawing.Point(96, 80);
			this.txtTiempo.MaxValue = 12;
			this.txtTiempo.MinValue = 0;
			this.txtTiempo.Name = "txtTiempo";
			this.txtTiempo.PromptChar = ' ';
			this.txtTiempo.Size = new System.Drawing.Size(48, 22);
			this.txtTiempo.TabIndex = 703;
			// 
			// btnGuardar
			// 
			this.btnGuardar.BackColor = System.Drawing.SystemColors.Control;
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(235, 120);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 704;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.BackColor = System.Drawing.SystemColors.Control;
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(312, 120);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 705;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(-480, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1264, 8);
			this.groupBox1.TabIndex = 706;
			this.groupBox1.TabStop = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(8, 56);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(67, 17);
			this.label5.TabIndex = 711;
			this.label5.Text = "Fecha Salida";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(8, 32);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(75, 17);
			this.label4.TabIndex = 709;
			this.label4.Text = "Fecha Ingreso";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaSalida
			// 
			appearance2.FontData.Name = "Tahoma";
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaSalida.Appearance = appearance2;
			this.dtFechaSalida.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaSalida.DateButtons.Add(dateButton1);
			this.dtFechaSalida.Enabled = false;
			this.dtFechaSalida.Format = "dd/MM/yyyy HH:mm";
			this.dtFechaSalida.Location = new System.Drawing.Point(96, 56);
			this.dtFechaSalida.Name = "dtFechaSalida";
			this.dtFechaSalida.NonAutoSizeHeight = 23;
			this.dtFechaSalida.Size = new System.Drawing.Size(152, 21);
			this.dtFechaSalida.SpinButtonsVisible = true;
			this.dtFechaSalida.TabIndex = 710;
			this.dtFechaSalida.Value = ((object)(resources.GetObject("dtFechaSalida.Value")));
			// 
			// dtFechaIngreso
			// 
			appearance3.FontData.Name = "Tahoma";
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaIngreso.Appearance = appearance3;
			this.dtFechaIngreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaIngreso.DateButtons.Add(dateButton2);
			this.dtFechaIngreso.Enabled = false;
			this.dtFechaIngreso.Format = "dd/MM/yyyy HH:mm";
			this.dtFechaIngreso.Location = new System.Drawing.Point(96, 32);
			this.dtFechaIngreso.Name = "dtFechaIngreso";
			this.dtFechaIngreso.NonAutoSizeHeight = 23;
			this.dtFechaIngreso.Size = new System.Drawing.Size(152, 21);
			this.dtFechaIngreso.SpinButtonsVisible = true;
			this.dtFechaIngreso.TabIndex = 708;
			this.dtFechaIngreso.Value = ((object)(resources.GetObject("dtFechaIngreso.Value")));
			// 
			// btnEditar
			// 
			this.btnEditar.BackColor = System.Drawing.SystemColors.Control;
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(159, 120);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 712;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// chkActivo
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkActivo.Appearance = appearance4;
			this.chkActivo.BackColor = System.Drawing.Color.Transparent;
			this.chkActivo.Enabled = false;
			this.chkActivo.Location = new System.Drawing.Point(248, 80);
			this.chkActivo.Name = "chkActivo";
			this.chkActivo.Size = new System.Drawing.Size(56, 22);
			this.chkActivo.TabIndex = 713;
			this.chkActivo.Text = "Activo";
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(8, 83);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(41, 17);
			this.label14.TabIndex = 702;
			this.label14.Text = "Tiempo";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(30, 17);
			this.label1.TabIndex = 714;
			this.label1.Text = "Local";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			appearance5.FontData.Name = "Tahoma";
			this.cmbBodega.Appearance = appearance5;
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 158;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 50;
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
			this.cmbBodega.Location = new System.Drawing.Point(96, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(208, 22);
			this.cmbBodega.TabIndex = 715;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// btnNuevo
			// 
			this.btnNuevo.BackColor = System.Drawing.SystemColors.Control;
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 120);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 716;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click_1);
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.SystemColors.Control;
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(80, 120);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 717;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// frmTiempoFacturas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(392, 150);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.chkActivo);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtFechaSalida);
			this.Controls.Add(this.dtFechaIngreso);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.txtTiempo);
			this.Controls.Add(this.label14);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmTiempoFacturas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Tiempo Estadia";
			this.Load += new System.EventHandler(this.frmTiempoFacturas_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		int iTiempoEstadia = 0;
		#endregion Variables

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmTiempoFacturas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "1125");

			if (!Funcion.Permiso("1753", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Configurar el Tiempo Estadia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}
			
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
			this.dtFechaIngreso.Value = System.DBNull.Value;
			this.dtFechaSalida.Value = System.DBNull.Value;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;

			if (this.dtFechaIngreso.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione Fecha Ingreso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaIngreso.Focus();
				return;
			}
			if (this.dtFechaSalida.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione Fecha Salida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaSalida.Focus();
				return;
			}
			if (Convert.ToInt32(this.txtTiempo.Value) == 0)
			{
				MessageBox.Show("El tiempo estadia debe ser mayor a CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtTiempo.Focus();
				return;
			}
			#endregion Validacion

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
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

					string sSQL = string.Format("Exec GrabaTiempoEstadia '{0}', '{1}', {2}, {3}, {4}, {5}", 
						Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd HH:mm"), Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd HH:mm"), 
						(int)this.txtTiempo.Value, iTiempoEstadia, this.chkActivo.Checked, (int)this.cmbBodega.Value);
					oCmdActualiza.CommandText = sSQL;
					oCmdActualiza.ExecuteNonQuery();

					oTransaction.Commit();
				}						
				catch (Exception ex)
				{
					oTransaction.Rollback();
					oConexion.Close();
					MessageBox.Show("Error", "Point Technoligy", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);												
					return;
				}
				finally
				{
					oConexion.Close();
				}
			}
			MessageBox.Show("Tiempo Estadia Registrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			this.btnCancelar_Click(sender, e);
			return;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.dtFechaIngreso.Enabled = false;
			this.dtFechaSalida.Enabled = false;
			this.txtTiempo.Enabled = false;
			this.chkActivo.Enabled = false;
			this.cmbBodega.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnCancelar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnEditar.Enabled = false;

			this.dtFechaIngreso.Value = System.DBNull.Value;
			this.dtFechaSalida.Value = System.DBNull.Value;
			this.txtTiempo.Value = 0;
			this.cmbBodega.Value = 0;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{

		}

		private void Consulta(int idTiempoEstadia)
		{
			string sSQL = string.Format("Exec ConsultaTiempoEstadia {0}", idTiempoEstadia);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{	
				iTiempoEstadia = dr.GetInt32(0);
				if (dr.GetValue(0) != System.DBNull.Value) this.dtFechaIngreso.Value = dr.GetDateTime(1);
				if (dr.GetValue(1) != System.DBNull.Value) this.dtFechaSalida.Value = dr.GetDateTime(2);
				this.txtTiempo.Value = dr.GetInt32(3);
				this.chkActivo.Checked = dr.GetBoolean(4);
				this.cmbBodega.Value = dr.GetInt32(5);
			}
			dr.Close();
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.dtFechaIngreso.Enabled = true;
			this.dtFechaSalida.Enabled = true;
			this.txtTiempo.Enabled = true;
			this.chkActivo.Enabled = true;
			this.cmbBodega.Enabled = true;

			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnConsultar.Enabled = false;
		}

		private void ultraCheckEditor2_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnNuevo_Click_1(object sender, System.EventArgs e)
		{
			this.dtFechaIngreso.Enabled = true;
			this.dtFechaSalida.Enabled = true;
			this.txtTiempo.Enabled = true;
			this.cmbBodega.Enabled = true;
			this.chkActivo.Enabled = true;
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

			iTiempoEstadia = 0;
			this.txtTiempo.Value = 0;
			this.dtFechaIngreso.Value = DateTime.Today;
			this.dtFechaSalida.Value = DateTime.Today;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			using (frmBuscaTiempoEstadia Busqueda = new frmBuscaTiempoEstadia())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.grdFacturas.ActiveRow.Cells["idTiempoModificacionEstadia"].Value);
				}
			}

			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			this.btnCancelar.Enabled = false;
			this.btnGuardar.Enabled = false;
		}
	}
}
