using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using Latinium.Services.OrdenDeCompra.Bodega;

namespace Latinium
{
	/// <summary>
	/// Summary description for ArticulosSinSerial.
	/// </summary>
	public class ArticulosSinSerial : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private C1.Data.C1DataSet cdsSerie;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		
		int IdArticulo = 0;
		string SCodigo = "";
		string SArticulo = "";
		int IdCompra;
		int IdDetCompra;
		int IDisponible = 3;
		int ICantidad = 0;
		int IBodega = 0;
		double DCantidad = 0; //
		string SNumero = "";
		double DRegistrados = 0;
		public int IdTipoFactura = 0;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnLlegaron;
		bool bUsado = false;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.GroupBox groupBox1;
		bool BEstado = false;
		string sNotas = "";

		public ArticulosSinSerial(double dCantidad, int idArticulo, string Codigo, string Articulo, int idCompra, string sNumero, int iDisponible, int idTipoFactura, int idDetCompra, bool bUso, int iBodega, bool bEstado, string SNotas)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			bUsado = bUso;
			IdDetCompra = idDetCompra;
			IdTipoFactura = idTipoFactura;
			IdArticulo = idArticulo;
			SArticulo = Articulo;
			SCodigo = Codigo;
			IdCompra = idCompra;
			SNumero = sNumero;
			IBodega = iBodega;
			BEstado = bEstado;

			if (iDisponible > 0)
			{
				IDisponible = 1;
				ICantidad = iDisponible;
			}
			if (iDisponible < 0)
			{
				IDisponible = 2;
				ICantidad = -iDisponible;
			}

			DCantidad = dCantidad;
			sNotas = SNotas;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.spnLlegaron = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.cdsSerie = new C1.Data.C1DataSet();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.spnLlegaron)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(8, 76);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(48, 15);
			this.ultraLabel4.TabIndex = 159;
			this.ultraLabel4.Text = "A Recibir";
			// 
			// spnLlegaron
			// 
			this.spnLlegaron.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnLlegaron.Location = new System.Drawing.Point(96, 72);
			this.spnLlegaron.MinValue = 0;
			this.spnLlegaron.Name = "spnLlegaron";
			this.spnLlegaron.PromptChar = ' ';
			this.spnLlegaron.Size = new System.Drawing.Size(144, 22);
			this.spnLlegaron.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnLlegaron.TabIndex = 158;
			this.spnLlegaron.Click += new System.EventHandler(this.spnLlegaron_Click);
			this.spnLlegaron.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.spnLlegaron_KeyPress);
			this.spnLlegaron.ValueChanged += new System.EventHandler(this.spnLlegaron_ValueChanged);
			this.spnLlegaron.Enter += new System.EventHandler(this.spnLlegaron_Enter);
			// 
			// ultraLabel3
			// 
			appearance1.ForeColor = System.Drawing.Color.Firebrick;
			this.ultraLabel3.Appearance = appearance1;
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.ultraLabel3.Location = new System.Drawing.Point(272, 78);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(0, 0);
			this.ultraLabel3.TabIndex = 157;
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(8, 12);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(38, 15);
			this.ultraLabel2.TabIndex = 156;
			this.ultraLabel2.Text = "Código";
			// 
			// txtCodigo
			// 
			appearance2.ForeColor = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance2;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(96, 8);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.ReadOnly = true;
			this.txtCodigo.Size = new System.Drawing.Size(392, 22);
			this.txtCodigo.TabIndex = 155;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(8, 44);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(42, 15);
			this.ultraLabel1.TabIndex = 154;
			this.ultraLabel1.Text = "Artículo";
			// 
			// txtArticulo
			// 
			appearance3.ForeColor = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance3;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Enabled = false;
			this.txtArticulo.Location = new System.Drawing.Point(96, 40);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.ReadOnly = true;
			this.txtArticulo.Size = new System.Drawing.Size(392, 22);
			this.txtArticulo.TabIndex = 153;
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(96, 200);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(80, 24);
			this.btCancelar.TabIndex = 162;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(8, 200);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(80, 24);
			this.btAceptar.TabIndex = 161;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// cdsSerie
			// 
			this.cdsSerie.BindingContextCtrl = this;
			this.cdsSerie.DataLibrary = "LibFacturacion";
			this.cdsSerie.DataLibraryUrl = "";
			this.cdsSerie.DataSetDef = "dsSerie";
			this.cdsSerie.FillOnRequest = false;
			this.cdsSerie.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSerie.Name = "cdsSerie";
			this.cdsSerie.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSerie.SchemaDef = null;
			this.cdsSerie.BeforeFill += new C1.Data.FillEventHandler(this.cdsSerie_BeforeFill);
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.Location = new System.Drawing.Point(8, 104);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(76, 15);
			this.ultraLabel5.TabIndex = 164;
			this.ultraLabel5.Text = "Observaciones";
			// 
			// txtObservaciones
			// 
			appearance4.ForeColor = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance4;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(96, 104);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(392, 72);
			this.txtObservaciones.TabIndex = 163;
			this.txtObservaciones.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtObservaciones_KeyPress);
			this.txtObservaciones.ValueChanged += new System.EventHandler(this.txtObservaciones_ValueChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 184);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(504, 8);
			this.groupBox1.TabIndex = 165;
			this.groupBox1.TabStop = false;
			// 
			// ArticulosSinSerial
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(658, 230);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.spnLlegaron);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtArticulo);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "ArticulosSinSerial";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ArticulosSinSerial_KeyDown);
			this.Load += new System.EventHandler(this.ArticulosSinSerial_Load);
			((System.ComponentModel.ISupportInitialize)(this.spnLlegaron)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		
		private void cdsSerie_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSerie.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		
		int idSerial = 0;
		private void ArticulosSinSerial_Load(object sender, System.EventArgs e)
		{
			if (BEstado) 
			{
				string sSQL = string.Format(
					"Select Bloquear From Compra Where idCompra = {0}",
					IdCompra
					);

				bool Bloqueado =
					Funcion.bEscalarSQL(cdsSerie, sSQL, true);

				if (Bloqueado)
				{
					this.spnLlegaron.Enabled = false;
					this.btAceptar.Enabled = false;
				}
			}

			this.txtArticulo.Text = SArticulo;
			this.txtCodigo.Text = SCodigo;

			// =====================================================
			// NUEVO PROCESO SOLO PARA ORDENES DE COMPRA
			// =====================================================

			if (IdTipoFactura == 2)
			{
				// OBTENER TOTAL REGISTRADO
				DRegistrados = Funcion.iEscalarSQL(
					cdsSerie,
					string.Format(
					@"Select ISNULL(SUM(Cantidad),0)
				  From Serie
				  Where idDetCompra = {0}
				  And idArticulo = {1}",
					IdDetCompra,
					IdArticulo
					)
					);

				// CALCULAR PENDIENTES
				double iPendientes = DCantidad - DRegistrados;

				if (iPendientes < 0)
					iPendientes = 0;

				// EL USUARIO SOLO INGRESA LO NUEVO
				this.spnLlegaron.Value = 0;

				// MOSTRAR INFORMACION
				this.ultraLabel3.Text =
					"Articulos Registrados " +
					DRegistrados +
					" de " +
					DCantidad +
					". Pendientes: " +
					iPendientes +
					".";
			}
			else
			{
				// =====================================================
				// PROCESO ORIGINAL
				// =====================================================

				idSerial = Funcion.iEscalarSQL(
					cdsSerie,
					string.Format(
					"Exec idSerial {0}",
					IdDetCompra
					)
					);

				if ((int)idSerial > 0)
				{
					DRegistrados = Funcion.iEscalarSQL(
						cdsSerie,
						string.Format(
						"Select Cantidad From Serie Where idSerie = {0}",
						idSerial
						)
						);
				}
				else
				{
					DRegistrados = 0;
				}

				this.spnLlegaron.Value = DRegistrados;

				this.ultraLabel3.Text =
					"Articulos Registrados " +
					DRegistrados +
					" de " +
					DCantidad +
					".";
			}

			//---------------------------------------------------------
			// SOLO VISUAL
			//---------------------------------------------------------
			this.ConfigurarModoSoloVisual();
		}


		//---------------------------------------------------------
		// CONFIGURA SOLO VISUAL
		//---------------------------------------------------------
		private void ConfigurarModoSoloVisual()
		{
			//-----------------------------------------------------
			// SOLO PARA ORDEN DE COMPRA
			//-----------------------------------------------------
			if (IdTipoFactura != 2)
				return;

			//-----------------------------------------------------
			// OBTENER ESTADO FACTURACION
			//-----------------------------------------------------
			int iEstadoFacturacion =
				Funcion.iEscalarSQL(
				cdsSerie,
				string.Format(
				@"SELECT ISNULL(Estado, 0)
		  FROM Compra
		  WHERE idCompra = {0}",
				IdCompra
				)
				);

			//-----------------------------------------------------
			// ESTADOS BLOQUEADOS
			// Solo estado 20 (factura en proceso) bloquea.
			// Estado 28 NO bloquea: permite re-recepcionar
			// mercaderia adicional antes de facturar (Escenario A)
			//-----------------------------------------------------
			if (iEstadoFacturacion == frmIngresoSeriales.EstadoPorFacturar)
			{
				this.spnLlegaron.Enabled = false;
				this.btAceptar.Enabled = false;

				this.ultraLabel3.Text +=
					" [SOLO LECTURA]";
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;			
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			int iNuevaCantidad = 0;

			#region Validacion

			if (IdTipoFactura == 2)
			{
				// =====================================================
				// NUEVA VALIDACION PARA ORDENES DE COMPRA
				// PERMITE REGISTROS PARCIALES ACUMULADOS
				// =====================================================

				// VALIDAR CANTIDAD MAYOR A 0
				if ((int)this.spnLlegaron.Value == 0)
				{
					MessageBox.Show(
						"Ingrese la cantidad recibida del Artículo",
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Error);

					return;
				}

				// OBTENER CANTIDAD YA REGISTRADA
				int iCantidadRegistrada = 0;

				string sSQLCantidad = string.Format(
					@"Select ISNULL(SUM(Cantidad),0)
	  From Serie
	  Where idDetCompra = {0}
	  And idArticulo = {1}",
					IdDetCompra,
					IdArticulo);

				DataSet dsCantidad = 	Funcion.dsEjecucion(MenuLatinium.sconexionEmpresa,sSQLCantidad);

				if (dsCantidad.Tables.Count > 0)
				{
					if (dsCantidad.Tables[0].Rows.Count > 0)
					{
						iCantidadRegistrada =
							Convert.ToInt32(dsCantidad.Tables[0].Rows[0][0]);
					}
				}

				// CALCULAR NUEVA CANTIDAD
				iNuevaCantidad =
					iCantidadRegistrada + (int)this.spnLlegaron.Value;

				// VALIDAR EXCESO
				if (iNuevaCantidad > DCantidad)
				{
					MessageBox.Show(
						string.Format(
						"No puede registrar ({0}) artículos.\n\n" +
						"Ya existen ({1}) registrados y la cantidad máxima es ({2}).",
						(int)this.spnLlegaron.Value,
						iCantidadRegistrada,
						DCantidad),
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Warning);

					return;
				}

				// VALIDAR PARCIAL
				if (iNuevaCantidad < DCantidad)
				{
					if (this.txtObservaciones.Text.ToString().Trim() == "")
					{
						MessageBox.Show(
							"El valor ingresado aún no completa la cantidad del documento.\n\nIngrese una OBSERVACIÓN.",
							"Point Technology",
							MessageBoxButtons.OK,
							MessageBoxIcon.Warning);

						return;
					}

					MessageBox.Show(
						"El artículo quedará parcialmente recibido.\n\n" +
						"Podrá registrar cantidades adicionales posteriormente hasta completar el total.",
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Warning);
				}
			}
			else
			{
				// =====================================================
				// VALIDACION ORIGINAL
				// =====================================================

				if ((int)this.spnLlegaron.Value == 0)
				{
					MessageBox.Show(
						"Ingrese la cantidad recibida del Artículo",
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Error);

					return;
				}

				if ((int)this.spnLlegaron.Value > DCantidad)
				{
					MessageBox.Show(
						string.Format(
						"El número de articulos ({0}) no puede ser mayor a ({1})",
						(int)this.spnLlegaron.Value,
						DCantidad),
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Error);

					return;
				}

				if ((int)this.spnLlegaron.Value < DCantidad)
				{
					if (this.txtObservaciones.Text.ToString().Trim() == "")
					{
						MessageBox.Show(
							"El valor que ingreso es menor al del Documento.\n\nIngrese una OBSERVACIÓN.",
							"Point Technology",
							MessageBoxButtons.OK,
							MessageBoxIcon.Warning);

						return;
					}

					MessageBox.Show(
						"El valor que ingreso es menor al del Documento.\n\nEste artículo estara en tránsito hasta que registre que todos los artículos han llegado.",
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Warning);
				}
			}

			#endregion Validacion
			using (SqlConnection oConexion = new SqlConnection(MenuLatinium.sconexionEmpresa))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;

				try
				{
					oTransaction = oConexion.BeginTransaction();

					SqlCommand oCmd = oConexion.CreateCommand();
					oCmd.Connection = oConexion;
					oCmd.Transaction = oTransaction;
					oCmd.CommandType = CommandType.Text;
					oCmd.CommandTimeout = 0;

					#region Variables

					if (idSerial == 0)
						idSerial = -1;

					DateTime dFecha = DateTime.Now;

					string sObservaciones =
						this.txtObservaciones.Text
						.ToString()
						.Replace("'", "''");

					#endregion Variables

					#region Graba Serie

					oCmd.CommandText = string.Format(
						@"Exec GrabaSerial 
					{0}, 
					{1}, 
					{2}, 
					'', 
					'', 
					{3}, 
					'{4}', 
					1, 
					0, 
					{5}, 
					{6}, 
					{7}, 
					'{8}', 
					'{9}'",
						idSerial,
						IdArticulo,
						IdCompra,
						IdTipoFactura,
						SNumero,
						IdDetCompra,
						IBodega,
						(int)this.spnLlegaron.Value,
						dFecha.ToString("yyyyMMdd"),
						sObservaciones);

					oCmd.ExecuteNonQuery();

					#endregion Graba Serie

					#region Guarda Observaciones

					if (sObservaciones.Length > 0)
					{
						oCmd.CommandText = string.Format(
							@"Exec ERMGuardaNotasArtSinSerial 
						{0}, 
						'{1}'",
							IdDetCompra,
							sObservaciones);

						oCmd.ExecuteNonQuery();
					}

					#endregion Guarda Observaciones

					#region Actualiza EstadoSerial

					if (IdTipoFactura == 2)
					{
//						// NUEVO PROCESO ACUMULATIVO
//						if (iNuevaCantidad == DCantidad)
//						{
//							oCmd.CommandText = string.Format(
//								@"Update DetCompra 
//			  Set EstadoSerial = 1 
//			  Where idDetCompra = {0}",
//								IdDetCompra);
//
//							oCmd.ExecuteNonQuery();
//						}
					}
					else
					{
						// PROCESO ORIGINAL
						if ((int)this.spnLlegaron.Value == DCantidad)
						{
							oCmd.CommandText = string.Format(
								@"Update DetCompra 
			  Set EstadoSerial = 1 
			  Where idDetCompra = {0}",
								IdDetCompra);

							oCmd.ExecuteNonQuery();
						}
					}

					#endregion Actualiza EstadoSerial

					#region Procesar Estado Orden Compra

					if (IdTipoFactura != 2 && IdTipoFactura != 4 && IdTipoFactura != 36)
					{
						oCmd.CommandText = string.Format(
							@"Exec EstadoOrdenDeCompra
					{0},
					{1}",
							IdCompra,
							IdTipoFactura);

						oCmd.ExecuteNonQuery();
					}

					#endregion Procesar Estado Orden Compra

					#region Actualizar Facturacion

					OrdenDeCompraService oService =
						new OrdenDeCompraService(cdsSerie);

					oService.ActualizarFacturacionDetalleNoSerial(
						oConexion,
						oTransaction,
						IdDetCompra,
						MenuLatinium.stUsuario);

					#endregion Actualizar Facturacion

					oTransaction.Commit();

					this.DialogResult = DialogResult.OK;
				}
				catch (Exception ex)
				{
					try
					{
						if (oTransaction != null)
							oTransaction.Rollback();
					}
					catch (Exception ex2)
					{
						MessageBox.Show(
							string.Format(
							"Rollback Exception Type: {0} {1}",
							ex2.GetType(),
							ex2.Message),
							"Point Technology",
							MessageBoxButtons.OK,
							MessageBoxIcon.Stop);

						return;
					}

					MessageBox.Show(
						string.Format(
						"Commit Exception Type : {0} {1}",
						ex.GetType(),
						ex.Message),
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Stop);
				}
				finally
				{
					oConexion.Close();
				}
			}
		}

		private void spnLlegaron_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{
				this.txtObservaciones.Enabled = true;
				this.txtObservaciones.Focus();
			}
		}

		private void spnLlegaron_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void spnLlegaron_Enter(object sender, System.EventArgs e)
		{
			this.spnLlegaron.SelectAll();
		}

		private void spnLlegaron_Click(object sender, System.EventArgs e)
		{
			this.spnLlegaron.SelectAll();
		}

		private void txtObservaciones_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtObservaciones_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{
				this.btAceptar.Enabled = true;
				this.btAceptar.Focus();
			}
		}

		private void ArticulosSinSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}

