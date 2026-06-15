using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ReservFact.
	/// </summary>
	public class ReservFact : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsCliente;
		private System.Windows.Forms.CheckBox chkCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ReservFact()
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
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.chkCliente = new System.Windows.Forms.CheckBox();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(48, 88);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(176, 23);
			this.btAceptar.TabIndex = 6;
			this.btAceptar.Text = "&Generar Cruce ";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cmbCliente);
			this.groupBox1.Controls.Add(this.chkCliente);
			this.groupBox1.Controls.Add(this.btAceptar);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(280, 144);
			this.groupBox1.TabIndex = 7;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Cruzar Reservas de Clientes Por Facturas";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.CaseSensitive = false;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// chkCliente
			// 
			this.chkCliente.Checked = true;
			this.chkCliente.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCliente.Location = new System.Drawing.Point(16, 32);
			this.chkCliente.Name = "chkCliente";
			this.chkCliente.TabIndex = 7;
			this.chkCliente.Text = "Todo Cliente";
			this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
			// 
			// cmbCliente
			// 
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DisplayMember = "";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Enabled = false;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(112, 32);
			this.cmbCliente.MaxDropDownItems = 30;
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(160, 21);
			this.cmbCliente.TabIndex = 127;
			this.cmbCliente.ValueMember = "";
			// 
			// ReservFact
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 165);
			this.Controls.Add(this.groupBox1);
			this.MaximizeBox = false;
			this.Name = "ReservFact";
			this.Text = "Reservas por Facturas";
			this.Load += new System.EventHandler(this.ReservFact_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ReservFact_Load(object sender, System.EventArgs e)
		{
			string stSelect = "Select idCliente, Codigo,Nombre from cliente where proveedor = 0 order by nombre";		

			cmbCliente.DataSource  = Funcion.dvProcedimiento(cdsCliente,stSelect);
			cmbCliente.DisplayMember = "Codigo";
			cmbCliente.ValueMember = "idCliente";
			cmbCliente.Rows.Band.Columns["idCliente"].Hidden=true;
		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void chkCliente_CheckedChanged(object sender, System.EventArgs e)
		{
			if (chkCliente.Checked == true)
				cmbCliente.Enabled = false;
					else
				cmbCliente.Enabled=true;
		
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string vIdCliente = "0";			
			if (chkCliente.Checked == false)
			{
				if (cmbCliente.Value == null) return;
				 vIdCliente = cmbCliente.Value.ToString();
			}
			
			string stProcedimiento = "exec P_ReservasClientes "+vIdCliente;
			try
			{
				Funcion.EjecutaSQL(cdsCliente,stProcedimiento);
			}
			catch(Exception ex)
			{
			MessageBox.Show(ex.Message);
			}
			Cursor = Cursors.Default;			
		}
	}
}
