using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolBloqueo.
	/// </summary>
	public class RolBloqueo : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optOpcion;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblTitulo;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private C1.Data.C1DataSet cdsRol;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolBloqueo()
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.optOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label2 = new System.Windows.Forms.Label();
			this.lblTitulo = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsRol = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.optOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRol)).BeginInit();
			this.SuspendLayout();
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(372, 276);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 11;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(132, 276);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 24);
			this.btAceptar.TabIndex = 10;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
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
			this.optOpcion.Location = new System.Drawing.Point(196, 224);
			this.optOpcion.Name = "optOpcion";
			this.optOpcion.Size = new System.Drawing.Size(176, 34);
			this.optOpcion.TabIndex = 9;
			this.optOpcion.Text = "Impedir";
			this.optOpcion.ValueChanged += new System.EventHandler(this.optOpcion_ValueChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(196, 190);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(38, 16);
			this.label2.TabIndex = 8;
			this.label2.Text = "Fecha:";
			// 
			// lblTitulo
			// 
			this.lblTitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitulo.ForeColor = System.Drawing.Color.Firebrick;
			this.lblTitulo.Location = new System.Drawing.Point(76, 43);
			this.lblTitulo.Name = "lblTitulo";
			this.lblTitulo.Size = new System.Drawing.Size(424, 129);
			this.lblTitulo.TabIndex = 7;
			this.lblTitulo.Text = "Esta opción IMPIDE que se realicen modificaciones en Roles del Mes Seleccionado";
			// 
			// dtFecha
			// 
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "MMM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(260, 190);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 6;
			this.dtFecha.Value = new System.DateTime(2006, 9, 1, 0, 0, 0, 0);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// cdsRol
			// 
			this.cdsRol.BindingContextCtrl = this;
			this.cdsRol.DataLibrary = "LibPersonal";
			this.cdsRol.DataLibraryUrl = "";
			this.cdsRol.DataSetDef = "dsPersonal";
			this.cdsRol.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRol.Name = "cdsRol";
			this.cdsRol.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRol.SchemaDef = null;
			this.cdsRol.BeforeFill += new C1.Data.FillEventHandler(this.cdsRol_BeforeFill);
			// 
			// RolBloqueo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(544, 316);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.optOpcion);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblTitulo);
			this.Controls.Add(this.dtFecha);
			this.Name = "RolBloqueo";
			this.Text = "RolBloqueo";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RolBloqueo_Closing);
			this.Load += new System.EventHandler(this.RolBloqueo_Load);
			((System.ComponentModel.ISupportInitialize)(this.optOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRol)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaLimite = (DateTime) this.dtFecha.Value;

			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea bloquear Roles de Pagos de "
				+ dtFechaLimite.ToString("MMM/yyyy") + "?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			// Caso de Impedir
			string stSelect = "UPDATE RolPago SET Bloqueado ";
			string stMensaje = "";
			if ((int) this.optOpcion.Value == 0)
			{
				stSelect += "= 1";
				stMensaje = "Asientos Bloqueados";
			}
			else
			{
				stSelect += "= 0";
				stMensaje = "Asientos Habilitados";
			}

			stSelect += " WHERE Mes = " + dtFechaLimite.Month.ToString() +
				" And Anio = " + dtFechaLimite.Year.ToString();
			Funcion.EjecutaSQL(cdsRol, stSelect, true);

			MessageBox.Show(stMensaje, "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
			this.Close();
		}

		private void cdsRol_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsRol.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void RolBloqueo_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Rol Bloqueo'";
			Funcion.EjecutaSQL(cdsRol, stAudita, true);
			dtFecha.Value = DateTime.Today.AddMonths(-1);;
		}

		private void RolBloqueo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Rol Bloqueo'";
			Funcion.EjecutaSQL(cdsRol, stAudita, true);
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void optOpcion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
