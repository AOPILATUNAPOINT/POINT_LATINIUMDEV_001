
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de MarceloReportes.
	/// </summary>
	public class MarceloReportes : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button button1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Button btVentasMensual;
		private System.Windows.Forms.Button btDepositos;
		private System.Windows.Forms.Button btBancos;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.Button btGenerarrAsientos;
		private C1.Data.C1DataSet cdsMarcelo;
		private System.Windows.Forms.Button button2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MarceloReportes()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.label1 = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btVentasMensual = new System.Windows.Forms.Button();
			this.btDepositos = new System.Windows.Forms.Button();
			this.btBancos = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.btGenerarrAsientos = new System.Windows.Forms.Button();
			this.cdsMarcelo = new C1.Data.C1DataSet();
			this.button2 = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarcelo)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(112, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(56, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Reportes";
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(8, 104);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(96, 23);
			this.button1.TabIndex = 1;
			this.button1.Text = "General Ventas";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// dtDesde
			// 
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MMM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(8, 56);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(120, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 3;
			this.dtDesde.Value = new System.DateTime(2008, 2, 21, 0, 0, 0, 0);
			// 
			// dtHasta
			// 
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MMM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(152, 56);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(120, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 4;
			this.dtHasta.Value = new System.DateTime(2008, 2, 21, 0, 0, 0, 0);
			// 
			// btVentasMensual
			// 
			this.btVentasMensual.Location = new System.Drawing.Point(8, 136);
			this.btVentasMensual.Name = "btVentasMensual";
			this.btVentasMensual.Size = new System.Drawing.Size(96, 23);
			this.btVentasMensual.TabIndex = 5;
			this.btVentasMensual.Text = "Ventas Mensual";
			this.btVentasMensual.Click += new System.EventHandler(this.btVentasMensual_Click);
			// 
			// btDepositos
			// 
			this.btDepositos.Location = new System.Drawing.Point(8, 168);
			this.btDepositos.Name = "btDepositos";
			this.btDepositos.Size = new System.Drawing.Size(96, 23);
			this.btDepositos.TabIndex = 6;
			this.btDepositos.Text = "Diario de Ventas";
			this.btDepositos.Click += new System.EventHandler(this.btDepositos_Click);
			// 
			// btBancos
			// 
			this.btBancos.Location = new System.Drawing.Point(8, 200);
			this.btBancos.Name = "btBancos";
			this.btBancos.Size = new System.Drawing.Size(96, 23);
			this.btBancos.TabIndex = 7;
			this.btBancos.Text = "Bancos";
			this.btBancos.Click += new System.EventHandler(this.btBancos_Click);
			// 
			// button3
			// 
			this.button3.Location = new System.Drawing.Point(136, 136);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(96, 23);
			this.button3.TabIndex = 9;
			this.button3.Text = "Arcaval";
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// btGenerarrAsientos
			// 
			this.btGenerarrAsientos.Location = new System.Drawing.Point(32, 232);
			this.btGenerarrAsientos.Name = "btGenerarrAsientos";
			this.btGenerarrAsientos.Size = new System.Drawing.Size(208, 23);
			this.btGenerarrAsientos.TabIndex = 10;
			this.btGenerarrAsientos.Text = "Generar Asientos Bancos";
			this.btGenerarrAsientos.Click += new System.EventHandler(this.btGenerarrAsientos_Click);
			// 
			// cdsMarcelo
			// 
			this.cdsMarcelo.BindingContextCtrl = this;
			this.cdsMarcelo.DataLibrary = "LibFacturacion";
			this.cdsMarcelo.DataLibraryUrl = "";
			this.cdsMarcelo.DataSetDef = "dsMarceloGeneralVentas";
			this.cdsMarcelo.FillOnRequest = false;
			this.cdsMarcelo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsMarcelo.Name = "cdsMarcelo";
			this.cdsMarcelo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMarcelo.SchemaDef = null;
			this.cdsMarcelo.BeforeFill += new C1.Data.FillEventHandler(this.cdsMarcelo_BeforeFill);
			// 
			// button2
			// 
			this.button2.Location = new System.Drawing.Point(32, 264);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(208, 23);
			this.button2.TabIndex = 11;
			this.button2.Text = "Generar Asientos Ventas";
			this.button2.Click += new System.EventHandler(this.button2_Click_1);
			// 
			// MarceloReportes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(280, 293);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.btGenerarrAsientos);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.btBancos);
			this.Controls.Add(this.btDepositos);
			this.Controls.Add(this.btVentasMensual);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.label1);
			this.Name = "MarceloReportes";
			this.Text = "Marcelo Reportes";
			this.Load += new System.EventHandler(this.MarceloReportes_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarcelo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
		
			this.Cursor = Cursors.WaitCursor;
		  int[,] mat= new int[7,4];
			mat[1,1] = 5;
			mat[1,2] = 11;
			mat[1,3] = 3;
			mat[2,1] = 18;
			mat[2,2] = 14;
			mat[2,3] = 19;
			mat[3,1] = 2;
			mat[3,2] = 17;
			mat[3,3] = 16;
			mat[4,1] = 6;
			mat[4,2] = 10;
			mat[4,3] = 9;
			mat[5,1] = 13;
			mat[5,2] = 12;
			mat[5,3] = 15;
			mat[6,1] = 1;
			mat[6,2] = 0;
			mat[6,3] = 0;

			DateTime dtFecha = (DateTime) dtHasta.Value;
			string fechita= dtFecha.ToString("yyyyMMdd");

			for (int i=1;i<7;i++)
			{
			
				ParameterFields paramFields = new ParameterFields ();

				ParameterField paramField = new ParameterField ();
				ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
				paramField.ParameterFieldName = "@bodega1";
				discreteVal.Value = mat[i,1]; //dtFechaIni.ToString("yyyy-MM-dd");
				paramField.CurrentValues.Add (discreteVal);
				paramFields.Add (paramField);

				ParameterField paramField1 = new ParameterField ();
				ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
				paramField1.ParameterFieldName = "@bodega2";
				discreteVal1.Value = mat[i,2];
				paramField1.CurrentValues.Add (discreteVal1);
				paramFields.Add (paramField1);

				ParameterField paramField2 = new ParameterField ();
				ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
				paramField2.ParameterFieldName = "@bodega3";
				discreteVal2.Value = mat[i,3];
				paramField2.CurrentValues.Add (discreteVal2);
				paramFields.Add (paramField2);

				ParameterField paramField4 = new ParameterField ();
				ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
				paramField4.ParameterFieldName = "@año";
				discreteVal4.Value = fechita.Substring(0,4); //"2008"; // 3 General 4 Resultados
				paramField4.CurrentValues.Add (discreteVal4);
				paramFields.Add (paramField4);
	
				ParameterField paramField5 = new ParameterField ();
				ParameterDiscreteValue discreteVal5 = new ParameterDiscreteValue ();
				paramField5.ParameterFieldName = "@mes";// "@mes";
				discreteVal5.Value = fechita.Substring(4,2); // 3 General 4 Resultados
				paramField5.CurrentValues.Add (discreteVal5);
				paramFields.Add (paramField5);

				ParameterField paramField6 = new ParameterField ();
				ParameterDiscreteValue discreteVal6= new ParameterDiscreteValue ();
				paramField6.ParameterFieldName = "@FECHAH";
				discreteVal6.Value = fechita; // 3 General 4 Resultados
				paramField6.CurrentValues.Add (discreteVal6);
				paramFields.Add (paramField6);

				Reporte miReporte;

				miReporte = new Reporte("GeneralVentas", "");
				miReporte.MdiParent = this.MdiParent;
				miReporte.crVista.ParameterFieldInfo = paramFields;
				miReporte.Show();

		
			}
				this.Cursor = Cursors.Default;	
		}

		private void MarceloReportes_Load(object sender, System.EventArgs e)
		{
			this.dtHasta.Value = DateTime.Today; 
			int dias= (int)DateTime.Today.Day; 
			this.dtDesde.Value=  DateTime.Today.AddDays(-dias+1);
		}

		private void btVentasMensual_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) dtHasta.Value;
			string fechita= dtFecha.ToString("yyyyMMdd");
			
			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@fecha";
			discreteVal.Value = fechita;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
			Reporte miReporte;
			miReporte = new Reporte("MarceloVentasMensual", "");
			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();
			this.Cursor = Cursors.Default;
		}

		private void btBancos_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			DateTime dtFechad = (DateTime) dtDesde.Value;
			DateTime dtFechah = (DateTime) dtHasta.Value;
			string fechad= dtFechad.ToString("yyyyMMdd");
			string fechah= dtFechah.ToString("yyyyMMdd");

			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@FECHAD2";
			discreteVal.Value = dtFechad.ToString("yyyyMMdd") ; //dtFechaIni.ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FECHAH2";
			discreteVal1.Value = dtFechah.ToString("yyyyMMdd");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);



			Reporte miReporte;

			miReporte = new Reporte("MarceloBanco", "");
			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();

			this.Cursor = Cursors.Default;

		}

		private void btDepositos_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			DateTime dtFechad = (DateTime) dtDesde.Value;
			DateTime dtFechah = (DateTime) dtHasta.Value;
			string fechad= dtFechad.ToString("yyyyMMdd");
			string fechah= dtFechah.ToString("yyyyMMdd");

			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@FECHAD";
			discreteVal.Value = dtFechad.ToString("yyyyMMdd") ; //dtFechaIni.ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FECHAH";
			discreteVal1.Value = dtFechah.ToString("yyyyMMdd");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);



			Reporte miReporte;

			miReporte = new Reporte("MarceloDeposito", "");
			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();

			this.Cursor = Cursors.Default;

		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) dtHasta.Value;
			string fechita= dtFecha.ToString("yyyyMMdd");
			
			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@fecha";
			discreteVal.Value = fechita;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
			Reporte miReporte;
			miReporte = new Reporte("MarceloVentasMensual2", "");
			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();
			this.Cursor = Cursors.Default;
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			int[,] mat= new int[7,4];
			mat[1,1] = 5;
			mat[1,2] = 11;
			mat[1,3] = 3;
			mat[2,1] = 18;
			mat[2,2] = 14;
			mat[2,3] = 2;
			mat[3,1] = 17;
			mat[3,2] = 16;
			mat[3,3] = 6;
			mat[4,1] = 10;
			mat[4,2] = 9;
			mat[4,3] = 13;
			mat[5,1] = 12;
			mat[5,2] = 15;
			mat[5,3] = 1;
			mat[6,1] = 8;
			mat[6,2] = 0;
			mat[6,3] = 0;

			

			DateTime dtFecha = (DateTime) dtHasta.Value;
			string fechita= dtFecha.ToString("yyyyMMdd");

			for (int i=1;i<7;i++)
			{
			
				ParameterFields paramFields = new ParameterFields ();

				ParameterField paramField = new ParameterField ();
				ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
				paramField.ParameterFieldName = "@bodega1";
				discreteVal.Value = mat[i,1]; //dtFechaIni.ToString("yyyy-MM-dd");
				paramField.CurrentValues.Add (discreteVal);
				paramFields.Add (paramField);

				ParameterField paramField1 = new ParameterField ();
				ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
				paramField1.ParameterFieldName = "@bodega2";
				discreteVal1.Value = mat[i,2];
				paramField1.CurrentValues.Add (discreteVal1);
				paramFields.Add (paramField1);

				ParameterField paramField2 = new ParameterField ();
				ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
				paramField2.ParameterFieldName = "@bodega3";
				discreteVal2.Value = mat[i,3];
				paramField2.CurrentValues.Add (discreteVal2);
				paramFields.Add (paramField2);

				ParameterField paramField4 = new ParameterField ();
				ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
				paramField4.ParameterFieldName = "@año";
				discreteVal4.Value = fechita.Substring(0,4); //"2008"; // 3 General 4 Resultados
				paramField4.CurrentValues.Add (discreteVal4);
				paramFields.Add (paramField4);
	
				ParameterField paramField5 = new ParameterField ();
				ParameterDiscreteValue discreteVal5 = new ParameterDiscreteValue ();
				paramField5.ParameterFieldName = "@mes";// "@mes";
				discreteVal5.Value = fechita.Substring(4,2); // 3 General 4 Resultados
				paramField5.CurrentValues.Add (discreteVal5);
				paramFields.Add (paramField5);

				ParameterField paramField6 = new ParameterField ();
				ParameterDiscreteValue discreteVal6= new ParameterDiscreteValue ();
				paramField6.ParameterFieldName = "@FECHAH";
				discreteVal6.Value = fechita; // 3 General 4 Resultados
				paramField6.CurrentValues.Add (discreteVal6);
				paramFields.Add (paramField6);

				Reporte miReporte;

				miReporte = new Reporte("GeneralVentas", "");
				miReporte.MdiParent = this.MdiParent;
				miReporte.crVista.ParameterFieldInfo = paramFields;
				miReporte.Show();
			}
			this.Cursor = Cursors.Default;	
		}

		private void btGenerarrAsientos_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechad = (DateTime) dtDesde.Value;
			DateTime dtFechah = (DateTime) dtHasta.Value;
			string fechad= dtFechad.ToString("yyyyMMdd");
			string fechah= dtFechah.ToString("yyyyMMdd");

   		string cadena= "exec marcelocreaasientoingreso "  + "'"+fechad.ToString() + "','"
			+fechad.ToString() + "','"+fechah.ToString() + "'";
	
			if (DialogResult.No == MessageBox.Show("Desea crear asientos de venta del "+ fechad.ToString() + " con fecha del "+ fechah.ToString() , "Informacion", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				return;
			}
		  Funcion.EjecutaSQL( cdsMarcelo ,cadena);
			MessageBox.Show("Asientos realizados con éxito");  
	
		}

		private void cdsMarcelo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMarcelo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void button2_Click_1(object sender, System.EventArgs e)
		{
			DateTime dtFechad = (DateTime) dtDesde.Value;
			DateTime dtFechah = (DateTime) dtHasta.Value;
			string fechad= dtFechad.ToString("yyyyMMdd");
			string fechah= dtFechah.ToString("yyyyMMdd");

			string cadena= "exec marcelocreaasientoingreso2 "  + "'"+fechad.ToString() + "','"
				+fechad.ToString() + "','"+fechah.ToString() + "'";
	
			if (DialogResult.No == MessageBox.Show("Desea crear asientos de venta del "+ fechad.ToString() + " con fecha del "+ fechah.ToString() , "Informacion", 
				MessageBoxButtons.YesNo, MessageBoxIcon.Question))
			{
				return;
			}
			Funcion.EjecutaSQL( cdsMarcelo ,cadena);
			MessageBox.Show("Asientos realizados con éxito");  

		}
	}
}
