using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de AsientoBloqueo.
	/// </summary>
	public class AsientoBloqueo : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optOpcion;
		private System.Windows.Forms.Label lblTitulo;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private System.Windows.Forms.HelpProvider helpProvider1;
		private C1.Data.C1DataSet cdsSeteoC;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public AsientoBloqueo()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblTitulo = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.optOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.helpProvider1 = new System.Windows.Forms.HelpProvider();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			this.SuspendLayout();
			// 
			// dtFecha
			// 
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MMM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(216, 152);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.TabIndex = 0;
			// 
			// lblTitulo
			// 
			this.lblTitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitulo.ForeColor = System.Drawing.Color.Firebrick;
			this.lblTitulo.Location = new System.Drawing.Point(32, 16);
			this.lblTitulo.Name = "lblTitulo";
			this.lblTitulo.Size = new System.Drawing.Size(424, 120);
			this.lblTitulo.TabIndex = 1;
			this.lblTitulo.Text = "Esta opción IMPIDE que se realicen modificaciones en asientos anteriores a la fec" +
				"ha indicada";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(152, 152);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(39, 16);
			this.label2.TabIndex = 2;
			this.label2.Text = "Fecha:";
			// 
			// optOpcion
			// 
			this.optOpcion.CheckedIndex = 0;
			this.optOpcion.ItemAppearance = appearance1;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Impedir";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Permitir";
			this.optOpcion.Items.Add(valueListItem1);
			this.optOpcion.Items.Add(valueListItem2);
			this.optOpcion.ItemSpacingHorizontal = 30;
			this.optOpcion.ItemSpacingVertical = 10;
			this.optOpcion.Location = new System.Drawing.Point(152, 184);
			this.optOpcion.Name = "optOpcion";
			this.optOpcion.Size = new System.Drawing.Size(176, 32);
			this.optOpcion.TabIndex = 3;
			this.optOpcion.Text = "Impedir";
			this.optOpcion.ValueChanged += new System.EventHandler(this.optOpcion_ValueChanged);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(88, 232);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(328, 232);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// helpProvider1
			// 
			this.helpProvider1.HelpNamespace = "C:\\Latinium\\Latinium.chm";
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			this.cdsSeteoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoC_BeforeFill);
			// 
			// AsientoBloqueo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(488, 285);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.optOpcion);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblTitulo);
			this.Controls.Add(this.dtFecha);
			this.helpProvider1.SetHelpKeyword(this, "Bloquear.htm");
			this.helpProvider1.SetHelpNavigator(this, System.Windows.Forms.HelpNavigator.Topic);
			this.MaximizeBox = false;
			this.Name = "AsientoBloqueo";
			this.helpProvider1.SetShowHelp(this, true);
			this.Text = "Bloqueo de Asientos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.AsientoBloqueo_Closing);
			this.Load += new System.EventHandler(this.AsientoBloqueo_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void AsientoBloqueo_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			this.dtFecha.Value = DateTime.Today;
			string stAudita = "Exec AuditaCrear 53, 6, 'Asi. Bloqueo'";
			Funcion.EjecutaSQL(cdsSeteoC, stAudita, true);
		}

		private void optOpcion_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int) this.optOpcion.Value == 0)
				this.lblTitulo.Text = "Esta opción IMPIDE que se realicen modificaciones en los asientos" 
					+ " que tengan una fecha anterior a la indicada";
			else
				this.lblTitulo.Text = "Esta opción PERMITE que se realicen modificaciones en los asientos" 
					+ " que tengan una fecha anterior a la indicada";
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			string stBloqueo = " Bloquear ";
			if ((int) this.optOpcion.Value == 1) stBloqueo = " DesBloquear ";
			if (DialogResult.No == MessageBox.Show(
        "¿Está seguro que desea" +  stBloqueo + "asientos anteriores a la fecha "
				+ this.dtFecha.Text.ToString() + "?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			// Caso de Impedir
			string stSelPago = "Update Pago Set Bloquear ";
			string stSelCompra = "Update Compra Set Bloquear ";
			string stSelect = "UPDATE Asiento SET Bloquear ";
			string stMensaje = "";
			if ((int) this.optOpcion.Value == 0)
			{
				stSelPago += " = 1 Where Bloquear = 0 ";
				stSelect += " = 1 Where Bloquear = 0 ";
				stSelCompra += " = 1 Where Bloquear = 0 ";
				stMensaje = "Asientos Bloqueados";
			}
			else
			{
				stSelPago += " = 0 Where Bloquear = 1 ";
				stSelect += "= 0 Where Bloquear = 1 ";
				stSelCompra += " = 0 Where Bloquear = 1 ";
				stMensaje = "Asientos Habilitados";
			}

			DateTime dtFechaLimite = (DateTime) this.dtFecha.Value;
			stSelPago += " And Fecha <= '" + dtFechaLimite.ToString("yyyyMMdd") + "'";
			stSelect += " And Fecha <= '" + dtFechaLimite.ToString("yyyyMMdd") + "'";
			stSelCompra += " And Fecha <= '" + dtFechaLimite.ToString("yyyyMMdd") + "'";

			Funcion.EjecutaSQL(cdsSeteoC, stSelPago, true);
			Funcion.EjecutaSQL(cdsSeteoC, stSelect, true);
			Funcion.EjecutaSQL(cdsSeteoC, stSelCompra, true);
			MessageBox.Show(stMensaje, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
			this.Close();
		}

		private void AsientoBloqueo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Asi. Bloqueo'";
			Funcion.EjecutaSQL(cdsSeteoC, stAudita, true);
		}

		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}
	}
}
