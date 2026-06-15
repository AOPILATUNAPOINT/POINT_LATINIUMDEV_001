using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Contometros.
	/// </summary>
	public class Contometros : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIni;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private C1.Data.C1DataSet cdsContometros;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDetalle;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdContometroDetalle;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtId;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Contometros()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_ContometrosDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idContometroDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoInicial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoFinal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Intermodulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desperdicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_Operativa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("F_Tecnica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idContometro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreBodega");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFechaIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsContometros = new C1.Data.C1DataSet();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.txtDetalle = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.grdContometroDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtId = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContometros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdContometroDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtId)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 16);
			this.label1.TabIndex = 18;
			this.label1.Text = "Fecha";
			// 
			// dtFechaIni
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaIni.Appearance = appearance1;
			this.dtFechaIni.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaIni.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContometros, "_Contometro.Fecha"));
			dateButton1.Caption = "Today";
			this.dtFechaIni.DateButtons.Add(dateButton1);
			this.dtFechaIni.Format = "dd/MMM/yyyy";
			this.dtFechaIni.Location = new System.Drawing.Point(72, 8);
			this.dtFechaIni.Name = "dtFechaIni";
			this.dtFechaIni.NonAutoSizeHeight = 23;
			this.dtFechaIni.Size = new System.Drawing.Size(120, 21);
			this.dtFechaIni.SpinButtonsVisible = true;
			this.dtFechaIni.TabIndex = 17;
			this.dtFechaIni.Value = new System.DateTime(2006, 3, 20, 0, 0, 0, 0);
			// 
			// cdsContometros
			// 
			this.cdsContometros.BindingContextCtrl = this;
			this.cdsContometros.DataLibrary = "LibFacturacion";
			this.cdsContometros.DataLibraryUrl = "";
			this.cdsContometros.DataSetDef = "dsContometros";
			this.cdsContometros.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsContometros.Name = "cdsContometros";
			this.cdsContometros.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsContometros.SchemaDef = null;
			this.cdsContometros.BeforeFill += new C1.Data.FillEventHandler(this.cdsContometros_BeforeFill);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Location = new System.Drawing.Point(216, 8);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(40, 14);
			this.ultraLabel1.TabIndex = 20;
			this.ultraLabel1.Text = "Detalle";
			// 
			// txtDetalle
			// 
			this.txtDetalle.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsContometros, "_Contometro.Detalle"));
			this.txtDetalle.Location = new System.Drawing.Point(264, 8);
			this.txtDetalle.Name = "txtDetalle";
			this.txtDetalle.Size = new System.Drawing.Size(360, 21);
			this.txtDetalle.TabIndex = 19;
			// 
			// grdContometroDetalle
			// 
			this.grdContometroDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdContometroDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdContometroDetalle.DataMember = "_Contometro.Contometro - ContometrosDetalle";
			this.grdContometroDetalle.DataSource = this.cdsContometros;
			appearance2.BackColor = System.Drawing.Color.White;
			this.grdContometroDetalle.DisplayLayout.Appearance = appearance2;
			this.grdContometroDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 67;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 45;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Width = 56;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Width = 56;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 56;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Width = 56;
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Width = 56;
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Width = 69;
			ultraGridColumn9.Header.VisiblePosition = 10;
			ultraGridColumn9.Width = 56;
			ultraGridColumn10.Header.VisiblePosition = 11;
			ultraGridColumn10.Width = 56;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 95;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 81;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.grdContometroDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdContometroDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdContometroDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdContometroDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grdContometroDetalle.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdContometroDetalle.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdContometroDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdContometroDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.grdContometroDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdContometroDetalle.Location = new System.Drawing.Point(24, 48);
			this.grdContometroDetalle.Name = "grdContometroDetalle";
			this.grdContometroDetalle.Size = new System.Drawing.Size(608, 272);
			this.grdContometroDetalle.TabIndex = 28;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_Contometro";
			this.barraDato1.DataNombreId = "idContometro";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsContometros;
			this.barraDato1.DataTabla = "Contometro";
			this.barraDato1.DataTablaHija = "ContometrosDetalle";
			this.barraDato1.Location = new System.Drawing.Point(16, 328);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 29;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca_1);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtId
			// 
			this.txtId.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtId.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsContometros, "_Contometro.idContometro"));
			this.txtId.Location = new System.Drawing.Point(536, 328);
			this.txtId.Name = "txtId";
			this.txtId.Size = new System.Drawing.Size(48, 21);
			this.txtId.TabIndex = 34;
			this.txtId.Text = "id";
			// 
			// Contometros
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(648, 365);
			this.Controls.Add(this.txtId);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.grdContometroDetalle);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtDetalle);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFechaIni);
			this.Name = "Contometros";
			this.Text = "Contometros";
			this.Load += new System.EventHandler(this.Contometros_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsContometros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdContometroDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtId)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
//		private bool bUsuario = false;
		private C1.Data.C1DataRow drContometro;
//		private Acceso miAcceso;
		int IdContometro = 0;

		private void Contometros_Load(object sender, System.EventArgs e)
		{
			txtId.Width = 0;
			#region Seguridad
//			miAcceso =new Acceso("0422", cdsContometros.StorageDataSet.Tables["Seguridad"]);
			#endregion
			if (IdContometro > 0)
			{
				BuscarAdicional(IdContometro);
			}
			else
			{
				this.barraDato1.ProximoId(4);
			}
			barraDato1.HabilitaControles(false);

		}
		private void BuscarAdicional(int idPersona)
		{
			if (idPersona > 0) 
			{
				this.barraDato1.IdRegistro = idPersona;
				this.barraDato1.ProximoId(5);
			}
		}
		private void cdsContometros_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsContometros.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			barraDato1.ProximoId(5);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (drContometro == null)
			{
				MessageBox.Show("No existe un registro para borrar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Está seguro que desea borrar el Registro?"
				+ ".\nEste proceso no puede ser revertido",	"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, 
				MessageBoxDefaultButton.Button2)) return;
			barraDato1.BorraRegistro();

		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
			try
			{
				drContometro = cdsContometros.TableViews["Contometro"].Rows[0];
			}
			catch{};
		}
		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			barraDato1.EditarRegistro();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			Funcion.TiempoInicia();
			
			barraDato1.CrearRegistro();
			drContometro["Fecha"] = DateTime.Today;
			barraDato1.ToolTipNuevo(Funcion.Tiempo("Crear"));
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			barraDato1.GrabaRegistro();
			string stExec = "Exec ContometroCarga " + txtId.Text;
			Funcion.EjecutaSQL(cdsContometros, stExec);
			barraDato1.ProximoId(5);
		}

		private void barraDato1_Refresca_1(object sender, System.EventArgs e)
		{
			Funcion.TiempoInicia();
			try
			{
				drContometro = cdsContometros.TableViews["Contometro"].Rows[0];
			}
			catch{};
			barraDato1.ToolTipBuscar(Funcion.Tiempo("Buscar"));
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{Contometro.idContometro} = " + txtId.Value.ToString();
			Reporte miRep = new Reporte("Contometros.rpt", stFiltro);
			miRep.MdiParent = MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}
	}
}
