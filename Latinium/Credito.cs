using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Credito.
	/// </summary>
	public class Credito : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label11;
		private C1.Data.C1DataSet cdsCredito;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbObservaciones;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCliente;
		private C1.Data.C1DataSet cdsCliente;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Credito()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		C1.Data.C1DataRow drCliente;

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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Estado", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.label13 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.cdsCredito = new C1.Data.C1DataSet();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.cmbObservaciones = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			((System.ComponentModel.ISupportInitialize)(this.cdsCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(48, 128);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(80, 16);
			this.label13.TabIndex = 102;
			this.label13.Text = "Observaciones";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(48, 104);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(39, 16);
			this.label11.TabIndex = 101;
			this.label11.Text = "Estado";
			// 
			// cdsCredito
			// 
			this.cdsCredito.BindingContextCtrl = this;
			this.cdsCredito.DataLibrary = "LibFacturacion";
			this.cdsCredito.DataLibraryUrl = "";
			this.cdsCredito.DataSetDef = "dsEstado";
			this.cdsCredito.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCredito.Name = "cdsCredito";
			this.cdsCredito.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCredito.SchemaDef = null;
			this.cdsCredito.BeforeFill += new C1.Data.FillEventHandler(this.cdsCredito_BeforeFill);
			// 
			// cmbEstado
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance1;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idEstado"));
			this.cmbEstado.DataMember = "Estado";
			this.cmbEstado.DataSource = this.cdsCredito;
			this.cmbEstado.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ArticuloGrupo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 58;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 99;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEstado.DisplayMember = "Estado";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(144, 104);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(176, 21);
			this.cmbEstado.TabIndex = 103;
			this.cmbEstado.ValueMember = "idEstado";
			this.cmbEstado.TextChanged += new System.EventHandler(this.cmbEstado_TextChanged);
			this.cmbEstado.RowSelected += new Infragistics.Win.UltraWinGrid.RowSelectedEventHandler(this.cmbEstado_RowSelected);
			this.cmbEstado.ValueChanged += new System.EventHandler(this.cmbEstado_ValueChanged);
			this.cmbEstado.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEstado_InitializeLayout);
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsCliente";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// cmbObservaciones
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbObservaciones.Appearance = appearance2;
			this.cmbObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbObservaciones.Cursor = System.Windows.Forms.Cursors.Hand;
			this.cmbObservaciones.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idObservaciones"));
			this.cmbObservaciones.DisplayLayout.AutoFitColumns = true;
			this.cmbObservaciones.DisplayMember = "Observaciones";
			this.cmbObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbObservaciones.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbObservaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbObservaciones.Location = new System.Drawing.Point(144, 128);
			this.cmbObservaciones.Name = "cmbObservaciones";
			this.cmbObservaciones.Size = new System.Drawing.Size(176, 21);
			this.cmbObservaciones.TabIndex = 108;
			this.cmbObservaciones.ValueMember = "idObservaciones";
			this.cmbObservaciones.ValueChanged += new System.EventHandler(this.cmbObservaciones_ValueChanged);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance3.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(40, 160);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(408, 176);
			this.ultraGrid1.TabIndex = 109;
			this.ultraGrid1.Text = "Detalle del estado de la solicitud de crédito de clientes";
			// 
			// cmbCliente
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCliente.Appearance = appearance8;
			this.cmbCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCliente.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCliente, "Cliente.idCliente"));
			this.cmbCliente.DisplayLayout.AutoFitColumns = true;
			this.cmbCliente.DisplayMember = "Nombre";
			this.cmbCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCliente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCliente.Location = new System.Drawing.Point(144, 72);
			this.cmbCliente.Name = "cmbCliente";
			this.cmbCliente.Size = new System.Drawing.Size(176, 21);
			this.cmbCliente.TabIndex = 111;
			this.cmbCliente.ValueMember = "idCliente";
			this.cmbCliente.ValueChanged += new System.EventHandler(this.cmbCliente_ValueChanged);
			this.cmbCliente.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCliente_InitializeLayout);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(48, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 110;
			this.label1.Text = "Cliente";
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Cliente";
			this.barraDato1.DataNombreId = "idCliente";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCliente;
			this.barraDato1.DataTabla = "Cliente";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(80, 8);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 112;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// Credito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(480, 342);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbCliente);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbObservaciones);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label11);
			this.Name = "Credito";
			this.Text = "Credito";
			this.Load += new System.EventHandler(this.Credito_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsCredito_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCredito.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cmbEstado_TextChanged(object sender, System.EventArgs e)
		{
			
		}

		private void cmbEstado_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void LlenaCombos()
		{
			string stSelect = string.Format("Exec ClienteCargaNombre 1");
			cmbCliente.DataSource  = Funcion.dvProcedimiento(cdsCredito, stSelect);
			cmbCliente.Rows.Band.Columns["idCliente"].Hidden = true;
			cmbCliente.Rows.Band.Columns["Codigo"].Hidden = true;

			stSelect = string.Format("Exec ListaClientesCredito");
			ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsCredito, stSelect);
		}

		private void Credito_Load(object sender, System.EventArgs e)
		{
			this.barraDato1.HabilitaControles(false);
			
			LlenaCombos();

			this.barraDato1.PosUltima();
		}

		private void cmbObservaciones_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void cmbEstado_ValueChanged(object sender, System.EventArgs e)
		{
				int Est = (int)cmbEstado.Value;
				string stSelect = string.Format("select idObservaciones, Observaciones from Observaciones where idEstado = {0}", Est);
				cmbObservaciones.DataSource  = Funcion.dvProcedimiento(cdsCredito, stSelect);
				cmbObservaciones.DisplayMember = "Observaciones";
				cmbObservaciones.ValueMember = "idObservaciones";
				cmbObservaciones.Rows.Band.Columns["idObservaciones"].Hidden = true;
		}

		private void cmbEstado_RowSelected(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
		{
			
		}

		private void cmbCliente_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbCliente_ValueChanged(object sender, System.EventArgs e)
		{
			if (barraDato1.bNuevo || barraDato1.bEditar)
			{
				string stSelect = string.Format("Exec ClientesCreditoEsrado {0}", (int)cmbCliente.Value);
				ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsCredito, stSelect);
			}
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			this.barraDato1.BorraRegistro();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drCliente = cdsCliente.TableViews["Casa"].Rows[0];
			}
			catch{};
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			this.barraDato1.GrabaRegistro();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
		}
	}
}
