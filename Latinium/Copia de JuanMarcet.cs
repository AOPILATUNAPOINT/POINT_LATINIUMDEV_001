using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de JuanMarcet.
	/// </summary>
	public class JuanMarcet : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechahasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaDesde;
		private Infragistics.Win.Misc.UltraButton btCuenta;
		private Infragistics.Win.Misc.UltraButton btTodo;
		private Infragistics.Win.Misc.UltraButton btAsiento;
		private Infragistics.Win.Misc.UltraButton btArticulos;
		private Infragistics.Win.Misc.UltraButton btClientes;
		private Infragistics.Win.Misc.UltraButton btTransacciones;
		private C1.Data.C1DataSet cdsCompra;

		DateTime dtFechaDesde;
		DateTime dtFechaHasta;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.Misc.UltraButton btPagos;

		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public JuanMarcet()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbFechahasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFechaDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btCuenta = new Infragistics.Win.Misc.UltraButton();
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.btTodo = new Infragistics.Win.Misc.UltraButton();
			this.btArticulos = new Infragistics.Win.Misc.UltraButton();
			this.btTransacciones = new Infragistics.Win.Misc.UltraButton();
			this.btClientes = new Infragistics.Win.Misc.UltraButton();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.btPagos = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechahasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(64, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(208, 40);
			this.label1.TabIndex = 51;
			this.label1.Text = "Importar Datos";
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(64, 104);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 16);
			this.label3.TabIndex = 195;
			this.label3.Text = "Hasta:";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(64, 80);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(40, 16);
			this.label5.TabIndex = 194;
			this.label5.Text = "Desde:";
			// 
			// cmbFechahasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechahasta.Appearance = appearance1;
			this.cmbFechahasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFechahasta.DateButtons.Add(dateButton1);
			this.cmbFechahasta.Format = "dd/MMM/yyyy";
			this.cmbFechahasta.Location = new System.Drawing.Point(112, 104);
			this.cmbFechahasta.Name = "cmbFechahasta";
			this.cmbFechahasta.NonAutoSizeHeight = 23;
			this.cmbFechahasta.Size = new System.Drawing.Size(144, 21);
			this.cmbFechahasta.SpinButtonsVisible = true;
			this.cmbFechahasta.TabIndex = 197;
			this.cmbFechahasta.Value = new System.DateTime(2012, 1, 24, 0, 0, 0, 0);
			// 
			// cmbFechaDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaDesde.Appearance = appearance2;
			this.cmbFechaDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbFechaDesde.DateButtons.Add(dateButton2);
			this.cmbFechaDesde.Format = "dd/MMM/yyyy";
			this.cmbFechaDesde.Location = new System.Drawing.Point(112, 80);
			this.cmbFechaDesde.Name = "cmbFechaDesde";
			this.cmbFechaDesde.NonAutoSizeHeight = 23;
			this.cmbFechaDesde.Size = new System.Drawing.Size(144, 21);
			this.cmbFechaDesde.SpinButtonsVisible = true;
			this.cmbFechaDesde.TabIndex = 196;
			this.cmbFechaDesde.Value = new System.DateTime(2012, 1, 24, 0, 0, 0, 0);
			// 
			// btCuenta
			// 
			this.btCuenta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCuenta.Location = new System.Drawing.Point(32, 184);
			this.btCuenta.Name = "btCuenta";
			this.btCuenta.Size = new System.Drawing.Size(72, 23);
			this.btCuenta.TabIndex = 201;
			this.btCuenta.Text = "Cuentas";
			this.btCuenta.Click += new System.EventHandler(this.btCuenta_Click);
			// 
			// btAsiento
			// 
			this.btAsiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAsiento.Location = new System.Drawing.Point(32, 216);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(72, 23);
			this.btAsiento.TabIndex = 202;
			this.btAsiento.Text = "Asientos";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// btTodo
			// 
			this.btTodo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btTodo.Location = new System.Drawing.Point(128, 136);
			this.btTodo.Name = "btTodo";
			this.btTodo.Size = new System.Drawing.Size(72, 23);
			this.btTodo.TabIndex = 203;
			this.btTodo.Text = "Todo";
			this.btTodo.Click += new System.EventHandler(this.btTodo_Click);
			// 
			// btArticulos
			// 
			this.btArticulos.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btArticulos.Location = new System.Drawing.Point(208, 184);
			this.btArticulos.Name = "btArticulos";
			this.btArticulos.Size = new System.Drawing.Size(96, 23);
			this.btArticulos.TabIndex = 204;
			this.btArticulos.Text = "Articulos";
			this.btArticulos.Click += new System.EventHandler(this.btArticulos_Click);
			// 
			// btTransacciones
			// 
			this.btTransacciones.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btTransacciones.Location = new System.Drawing.Point(208, 248);
			this.btTransacciones.Name = "btTransacciones";
			this.btTransacciones.Size = new System.Drawing.Size(96, 23);
			this.btTransacciones.TabIndex = 205;
			this.btTransacciones.Text = "Transacciones";
			this.btTransacciones.Click += new System.EventHandler(this.btTransacciones_Click);
			// 
			// btClientes
			// 
			this.btClientes.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btClientes.Location = new System.Drawing.Point(208, 216);
			this.btClientes.Name = "btClientes";
			this.btClientes.Size = new System.Drawing.Size(96, 23);
			this.btClientes.TabIndex = 206;
			this.btClientes.Text = "Clientes - Prov.";
			this.btClientes.Click += new System.EventHandler(this.btClientes_Click);
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
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
			// btPagos
			// 
			this.btPagos.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPagos.Location = new System.Drawing.Point(208, 280);
			this.btPagos.Name = "btPagos";
			this.btPagos.Size = new System.Drawing.Size(96, 23);
			this.btPagos.TabIndex = 207;
			this.btPagos.Text = "Pagos";
			this.btPagos.Click += new System.EventHandler(this.btPagos_Click);
			// 
			// JuanMarcet
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(360, 333);
			this.Controls.Add(this.btPagos);
			this.Controls.Add(this.btClientes);
			this.Controls.Add(this.btTransacciones);
			this.Controls.Add(this.btArticulos);
			this.Controls.Add(this.btTodo);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btCuenta);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.cmbFechahasta);
			this.Controls.Add(this.cmbFechaDesde);
			this.Controls.Add(this.label1);
			this.Name = "JuanMarcet";
			this.Text = "Juan Marcet";
			this.Load += new System.EventHandler(this.JuanMarcet_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFechahasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void JuanMarcet_Load(object sender, System.EventArgs e)
		{
			cmbFechaDesde.Value =DateTime.Today;
			cmbFechahasta.Value =DateTime.Today;

			dtFechaDesde = (DateTime) cmbFechaDesde.Value;
			dtFechaHasta = (DateTime) cmbFechahasta.Value;
		}

		private void btArticulos_Click(object sender, System.EventArgs e)
		{
			dtFechaDesde = (DateTime) cmbFechaDesde.Value;
			dtFechaHasta = (DateTime) cmbFechahasta.Value;
			
			
			if (MessageBox.Show("¿Desea Subir informacion de articulos?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			this.Cursor = Cursors.WaitCursor;
			string stExec = "Exec JuanMarcetArticulo '"+ dtFechaDesde.ToString("yyyyMMdd")  + "','"
				+ dtFechaHasta.ToString("yyyyMMdd")  + "'"  ;
			Funcion.EjecutaSQL(cdsSeteo , stExec, true);
			this.Cursor = Cursors.Default;
			MessageBox.Show("Informacion Subida");  

		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btTodo_Click(object sender, System.EventArgs e)
		{
			dtFechaDesde = (DateTime) cmbFechaDesde.Value;
			dtFechaHasta = (DateTime) cmbFechahasta.Value;
			
			if (MessageBox.Show("¿Desea Subir informacion del sistema de Oracle?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			this.Cursor = Cursors.WaitCursor;
			string stExec = "Exec JuanMarcetTodo '"+ dtFechaDesde.ToString("yyyyMMdd")  + "','"
				+ dtFechaHasta.ToString("yyyyMMdd")  + "'"  ;
			Funcion.EjecutaSQL(cdsSeteo , stExec, true);
			this.Cursor = Cursors.Default;
			MessageBox.Show("Informacion Subida");  
		}

		private void btCuenta_Click(object sender, System.EventArgs e)
		{
			dtFechaDesde = (DateTime) cmbFechaDesde.Value;
			dtFechaHasta = (DateTime) cmbFechahasta.Value;
			
			if (MessageBox.Show("¿Desea Subir plan de cuentas?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			this.Cursor = Cursors.WaitCursor;
			string stExec = "Exec JuanMarcetCuenta '"+ dtFechaDesde.ToString("yyyyMMdd")  + "','"
				+ dtFechaHasta.ToString("yyyyMMdd")  + "'"  ;
			Funcion.EjecutaSQL(cdsSeteo , stExec, true);
			this.Cursor = Cursors.Default;
			MessageBox.Show("Informacion Subida");  
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			dtFechaDesde = (DateTime) cmbFechaDesde.Value;
			dtFechaHasta = (DateTime) cmbFechahasta.Value;
			
			if (MessageBox.Show("¿Desea Subir informacion de asientos?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			this.Cursor = Cursors.WaitCursor;
			string stExec = "Exec JuanMarcetAsiento '"+ dtFechaDesde.ToString("yyyyMMdd")  + "','"
				+ dtFechaHasta.ToString("yyyyMMdd")  + "'"  ;
			Funcion.EjecutaSQL(cdsSeteo , stExec, true);
			this.Cursor = Cursors.Default;
			MessageBox.Show("Informacion Subida");  
		}

		private void btClientes_Click(object sender, System.EventArgs e)
		{
			dtFechaDesde = (DateTime) cmbFechaDesde.Value;
			dtFechaHasta = (DateTime) cmbFechahasta.Value;
			
			if (MessageBox.Show("¿Desea Subir Clientes y Proveedores?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			this.Cursor = Cursors.WaitCursor;
			string stExec = "Exec JuanMarcetCLientes '"+ dtFechaDesde.ToString("yyyyMMdd")  + "','"
				+ dtFechaHasta.ToString("yyyyMMdd")  + "'"  ;
			Funcion.EjecutaSQL(cdsSeteo , stExec, true);
			this.Cursor = Cursors.Default;
			MessageBox.Show("Informacion Subida");  		}

		private void btTransacciones_Click(object sender, System.EventArgs e)
		{
			dtFechaDesde = (DateTime) cmbFechaDesde.Value;
			dtFechaHasta = (DateTime) cmbFechahasta.Value;
			
			if (MessageBox.Show("¿Desea Subir transacciones de facturación?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			this.Cursor = Cursors.WaitCursor;
			string stExec = "Exec JuanMarcetTransacciones '"+ dtFechaDesde.ToString("yyyyMMdd")  + "','"
				+ dtFechaHasta.ToString("yyyyMMdd")  + "'"  ;
			Funcion.EjecutaSQL(cdsSeteo , stExec, true);
			this.Cursor = Cursors.Default;
			MessageBox.Show("Informacion Subida");  	
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btBorrar_Click(object sender, System.EventArgs e)
		{
			dtFechaDesde = (DateTime) cmbFechaDesde.Value;
			dtFechaHasta = (DateTime) cmbFechahasta.Value;
			
			if (MessageBox.Show("¿Desea borrar asientos vacíos?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			this.Cursor = Cursors.WaitCursor;
			string stExec = "Exec JuanMarcetBorrar '"+ dtFechaDesde.ToString("yyyyMMdd")  + "','"
				+ dtFechaHasta.ToString("yyyyMMdd")  + "'"  ;
			Funcion.EjecutaSQL(cdsSeteo , stExec, true);
			this.Cursor = Cursors.Default;
			MessageBox.Show("Informacion Subida");  	
		
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btPagos_Click(object sender, System.EventArgs e)
		{
			  
			dtFechaDesde = (DateTime) cmbFechaDesde.Value;
			dtFechaHasta = (DateTime) cmbFechahasta.Value;
			
			if (MessageBox.Show("¿Desea Subir transacciones de facturación?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			this.Cursor = Cursors.WaitCursor;
			string stExec = "Exec JuanMarcetPagos '"+ dtFechaDesde.ToString("yyyyMMdd")  + "','"
				+ dtFechaHasta.ToString("yyyyMMdd")  + "'"  ;
			Funcion.EjecutaSQL(cdsSeteo , stExec, true);
			this.Cursor = Cursors.Default;
			MessageBox.Show("Informacion Subida");  
			
		}
	}
}
