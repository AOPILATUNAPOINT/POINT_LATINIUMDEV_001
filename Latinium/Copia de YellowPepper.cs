using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Globalization;
using System.Security.Cryptography;  
using System.Text;  

namespace Latinium
{
	/// <summary>
	/// Descripción breve de YellowPepper.
	/// </summary>
	public class YellowPepper : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btArchivo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		private Infragistics.Win.Misc.UltraButton btSubir;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVerifica;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.Misc.UltraButton btDesencriptar;
		private Infragistics.Win.Misc.UltraButton btSubirArchivo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public YellowPepper()
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
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btArchivo = new Infragistics.Win.Misc.UltraButton();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btSubir = new Infragistics.Win.Misc.UltraButton();
			this.chkVerifica = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.btDesencriptar = new Infragistics.Win.Misc.UltraButton();
			this.btSubirArchivo = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// btCliente
			// 
			this.btCliente.Location = new System.Drawing.Point(32, 16);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(72, 32);
			this.btCliente.TabIndex = 1;
			this.btCliente.Text = "Leer";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.Location = new System.Drawing.Point(24, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(584, 232);
			this.ultraGrid1.TabIndex = 2;
			// 
			// btArchivo
			// 
			this.btArchivo.Location = new System.Drawing.Point(584, 16);
			this.btArchivo.Name = "btArchivo";
			this.btArchivo.Size = new System.Drawing.Size(24, 24);
			this.btArchivo.TabIndex = 14;
			this.btArchivo.Text = "...";
			this.btArchivo.Click += new System.EventHandler(this.btArchivo_Click);
			// 
			// txtPago
			// 
			this.txtPago.Location = new System.Drawing.Point(432, 16);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(144, 21);
			this.txtPago.TabIndex = 13;
			this.txtPago.Text = "c:\\Yellow\\";
			this.txtPago.ValueChanged += new System.EventHandler(this.txtPago_ValueChanged);
			// 
			// btSubir
			// 
			this.btSubir.Location = new System.Drawing.Point(120, 16);
			this.btSubir.Name = "btSubir";
			this.btSubir.Size = new System.Drawing.Size(72, 32);
			this.btSubir.TabIndex = 15;
			this.btSubir.Text = "Subir  N.";
			this.btSubir.Click += new System.EventHandler(this.btSubir_Click);
			// 
			// chkVerifica
			// 
			this.chkVerifica.Checked = true;
			this.chkVerifica.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkVerifica.Location = new System.Drawing.Point(368, 16);
			this.chkVerifica.Name = "chkVerifica";
			this.chkVerifica.Size = new System.Drawing.Size(64, 20);
			this.chkVerifica.TabIndex = 16;
			this.chkVerifica.Text = "Verifica";
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsSeteoF";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// btDesencriptar
			// 
			this.btDesencriptar.Location = new System.Drawing.Point(200, 16);
			this.btDesencriptar.Name = "btDesencriptar";
			this.btDesencriptar.Size = new System.Drawing.Size(72, 32);
			this.btDesencriptar.TabIndex = 17;
			this.btDesencriptar.Text = "Encriptar";
			this.btDesencriptar.Visible = false;
			this.btDesencriptar.Click += new System.EventHandler(this.btDesencriptar_Click);
			this.btDesencriptar.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btDesencriptar_MouseUp);
			// 
			// btSubirArchivo
			// 
			this.btSubirArchivo.Location = new System.Drawing.Point(280, 16);
			this.btSubirArchivo.Name = "btSubirArchivo";
			this.btSubirArchivo.Size = new System.Drawing.Size(72, 32);
			this.btSubirArchivo.TabIndex = 18;
			this.btSubirArchivo.Text = "Subir  Enc.";
			this.btSubirArchivo.Click += new System.EventHandler(this.btSubirArchivo_Click);
			// 
			// YellowPepper
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(632, 318);
			this.Controls.Add(this.btSubirArchivo);
			this.Controls.Add(this.btDesencriptar);
			this.Controls.Add(this.chkVerifica);
			this.Controls.Add(this.btSubir);
			this.Controls.Add(this.btArchivo);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btCliente);
			this.Name = "YellowPepper";
			this.Text = "YellowPepper";
			this.Load += new System.EventHandler(this.YellowPepper_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Decrypts specified ciphertext using Rijndael symmetric key algorithm.
		/// </summary>
		/// <param name="cipherText">
		/// Base64-formatted ciphertext value.
		/// </param>
		/// <param name="passPhrase">
		/// Passphrase from which a pseudo-random password will be derived. The
		/// derived password will be used to generate the encryption key.
		/// Passphrase can be any string. In this example we assume that this
		/// passphrase is an ASCII string.
		/// </param>
		/// <param name="saltValue">
		/// Salt value used along with passphrase to generate password. Salt can
		/// be any string. In this example we assume that salt is an ASCII string.
		/// </param>
		/// <param name="hashAlgorithm">
		/// Hash algorithm used to generate password. Allowed values are: "MD5" and
		/// "SHA1". SHA1 hashes are a bit slower, but more secure than MD5 hashes.
		/// </param>
		/// <param name="passwordIterations">
		/// Number of iterations used to generate password. One or two iterations
		/// should be enough.
		/// </param>
		/// <param name="initVector">
		/// Initialization vector (or IV). This value is required to encrypt the
		/// first block of plaintext data. For RijndaelManaged class IV must be
		/// exactly 16 ASCII characters long.
		/// </param>
		/// <param name="keySize">
		/// Size of encryption key in bits. Allowed values are: 128, 192, and 256.
		/// Longer keys are more secure than shorter keys.
		/// </param>
		/// <returns>
		/// Decrypted string value.
		/// </returns>
		/// <remarks>
		/// Most of the logic in this function is similar to the Encrypt
		/// logic. In order for decryption to work, all parameters of this function
		/// - except cipherText value - must match the corresponding parameters of
		/// the Encrypt function which was called to generate the
		/// ciphertext.
		/// </remarks>
		public static string Decrypt(string   cipherText,
			string   passPhrase,
			string   saltValue,
			string   hashAlgorithm,
			int      passwordIterations,
			string   initVector,
			int      keySize)
		{
			// Convert strings defining encryption key characteristics into byte
			// arrays. Let us assume that strings only contain ASCII codes.
			// If strings include Unicode characters, use Unicode, UTF7, or UTF8
			// encoding.
			byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
			byte[] saltValueBytes  = Encoding.ASCII.GetBytes(saltValue);

			// Convert our ciphertext into a byte array.
			byte[] cipherTextBytes = Convert.FromBase64String(cipherText);
        
			// First, we must create a password, from which the key will be 
			// derived. This password will be generated from the specified 
			// passphrase and salt value. The password will be created using
			// the specified hash algorithm. Password creation can be done in
			// several iterations.
			PasswordDeriveBytes password = new PasswordDeriveBytes(
				passPhrase, 
				saltValueBytes, 
				hashAlgorithm, 
				passwordIterations);
        
			// Use the password to generate pseudo-random bytes for the encryption
			// key. Specify the size of the key in bytes (instead of bits).
			byte[] keyBytes = password.GetBytes(keySize / 8);
        
			// Create uninitialized Rijndael encryption object.
			RijndaelManaged    symmetricKey = new RijndaelManaged();
        
			// It is reasonable to set encryption mode to Cipher Block Chaining
			// (CBC). Use default options for other symmetric key parameters.
			symmetricKey.Mode = CipherMode.CBC;
        
			// Generate decryptor from the existing key bytes and initialization 
			// vector. Key size will be defined based on the number of the key 
			// bytes.
			ICryptoTransform decryptor = symmetricKey.CreateDecryptor(
				keyBytes, 
				initVectorBytes);
        
			// Define memory stream which will be used to hold encrypted data.
			MemoryStream  memoryStream = new MemoryStream(cipherTextBytes);
                
			// Define cryptographic stream (always use Read mode for encryption).
			CryptoStream  cryptoStream = new CryptoStream(memoryStream, 
				decryptor,
				CryptoStreamMode.Read);

			// Since at this point we don't know what the size of decrypted data
			// will be, allocate the buffer long enough to hold ciphertext;
			// plaintext is never longer than ciphertext.
			byte[] plainTextBytes = new byte[cipherTextBytes.Length];
        
			// Start decrypting.
			int decryptedByteCount = cryptoStream.Read(plainTextBytes, 
				0, plainTextBytes.Length);
     
			// Close both streams.
			memoryStream.Close();
			cryptoStream.Close();
      
			// Convert decrypted data into a string. 
			// Let us assume that the original plaintext string was UTF8-encoded.
			string plainText = Encoding.UTF8.GetString(plainTextBytes, 
				0, decryptedByteCount);

			// Return decrypted string.   
			return plainText;
		}
	
		/// <summary>
		/// Encrypts specified plaintext using Rijndael symmetric key algorithm
		/// and returns a base64-encoded result.
		/// </summary>
		/// <param name="plainText">
		/// Plaintext value to be encrypted.
		/// </param>
		/// <param name="passPhrase">
		/// Passphrase from which a pseudo-random password will be derived. The
		/// derived password will be used to generate the encryption key.
		/// Passphrase can be any string. In this example we assume that this
		/// passphrase is an ASCII string.
		/// </param>
		/// <param name="saltValue">
		/// Salt value used along with passphrase to generate password. Salt can
		/// be any string. In this example we assume that salt is an ASCII string.
		/// </param>
		/// <param name="hashAlgorithm">
		/// Hash algorithm used to generate password. Allowed values are: "MD5" and
		/// "SHA1". SHA1 hashes are a bit slower, but more secure than MD5 hashes.
		/// </param>
		/// <param name="passwordIterations">
		/// Number of iterations used to generate password. One or two iterations
		/// should be enough.
		/// </param>
		/// <param name="initVector">
		/// Initialization vector (or IV). This value is required to encrypt the
		/// first block of plaintext data. For RijndaelManaged class IV must be 
		/// exactly 16 ASCII characters long.
		/// </param>
		/// <param name="keySize">
		/// Size of encryption key in bits. Allowed values are: 128, 192, and 256. 
		/// Longer keys are more secure than shorter keys.
		/// </param>
		/// <returns>
		/// Encrypted value formatted as a base64-encoded string.
		/// </returns>
		public static string Encrypt(string   plainText,
			string   passPhrase,
			string   saltValue,
			string   hashAlgorithm,
			int      passwordIterations,
			string   initVector,
			int      keySize)
		{
			// Convert strings into byte arrays.
			// Let us assume that strings only contain ASCII codes.
			// If strings include Unicode characters, use Unicode, UTF7, or UTF8 
			// encoding.
			byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
			byte[] saltValueBytes  = Encoding.ASCII.GetBytes(saltValue);
        
			// Convert our plaintext into a byte array.
			// Let us assume that plaintext contains UTF8-encoded characters.
			byte[] plainTextBytes  = Encoding.UTF8.GetBytes(plainText);
        
			// First, we must create a password, from which the key will be derived.
			// This password will be generated from the specified passphrase and 
			// salt value. The password will be created using the specified hash 
			// algorithm. Password creation can be done in several iterations.
			PasswordDeriveBytes password = new PasswordDeriveBytes(
				passPhrase, 
				saltValueBytes, 
				hashAlgorithm, 
				passwordIterations);
        
			// Use the password to generate pseudo-random bytes for the encryption
			// key. Specify the size of the key in bytes (instead of bits).
			byte[] keyBytes = password.GetBytes(keySize / 8);
 
			// Create uninitialized Rijndael encryption object.
			RijndaelManaged symmetricKey = new RijndaelManaged();

			// It is reasonable to set encryption mode to Cipher Block Chaining
			// (CBC). Use default options for other symmetric key parameters.
			symmetricKey.Mode = CipherMode.CBC;        
        
			// Generate encryptor from the existing key bytes and initialization 
			// vector. Key size will be defined based on the number of the key 
			// bytes.
			ICryptoTransform encryptor = symmetricKey.CreateEncryptor(
				keyBytes, 
				initVectorBytes);
        
			// Define memory stream which will be used to hold encrypted data.
			MemoryStream memoryStream = new MemoryStream();        
                
			// Define cryptographic stream (always use Write mode for encryption).
			CryptoStream cryptoStream = new CryptoStream(memoryStream, 
				encryptor,
				CryptoStreamMode.Write);
			// Start encrypting.
			cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                
			// Finish encrypting.
			cryptoStream.FlushFinalBlock();

			// Convert our encrypted data from a memory stream into a byte array.
			byte[] cipherTextBytes = memoryStream.ToArray();
                
			// Close both streams.
			memoryStream.Close();
			cryptoStream.Close();
        
			// Convert encrypted data into a base64-encoded string.
			string cipherText = Convert.ToBase64String(cipherTextBytes);

			// Return encrypted string.
			return cipherText;
		}

		public static string DecryptWith3DES(string cipherText,
			string passPhrase,
			string saltValue,
			string hashAlgorithm,
			int passwordIterations,
			string initVector,
			int keySize)
		{
			byte[] key, iv;

			ComposePasswordAndIV(passPhrase, saltValue, hashAlgorithm, passwordIterations, initVector, keySize, out key, out iv);

			byte[] cipherBytes = System.Convert.FromBase64String(cipherText);

			// Create a new MemoryStream using the passed 
			// array of encrypted data.
			MemoryStream msDecrypt = new MemoryStream(cipherBytes);

			// Create a CryptoStream using the MemoryStream 
			// and the passed key and initialization vector (IV).
			CryptoStream csDecrypt = new CryptoStream(msDecrypt,
				new TripleDESCryptoServiceProvider().CreateDecryptor(key, iv),
				CryptoStreamMode.Read);

			// Create buffer to hold the decrypted data.
			byte[] plainBytes = new byte[cipherBytes.Length];

			// Read the decrypted data out of the crypto stream
			// and place it into the temporary buffer.
			csDecrypt.Read(plainBytes, 0, plainBytes.Length);

			//Convert the buffer into a string and return it.
			return new ASCIIEncoding().GetString(plainBytes).TrimEnd('\0');
		}

		private static void ComposePasswordAndIV(string passPhrase,
			string saltValue,
			string hashAlgorithm,
			int passwordIterations,
			string initVector,
			int keySize,
			out byte[] key,
			out byte[] iv)
		{
			// Convert strings into byte arrays.
			// Let us assume that strings only contain ASCII codes.
			// If strings include Unicode characters, use Unicode, UTF7, or UTF8 
			// encoding.
			iv = Encoding.ASCII.GetBytes(initVector);
			byte[] saltValueBytes = Encoding.ASCII.GetBytes(saltValue);

			// First, we must create a password, from which the key will be derived.
			// This password will be generated from the specified passphrase and
			// salt value. The password will be created using the specified hash
			// algorithm. Password creation can be done in several iterations.
			PasswordDeriveBytes password = new PasswordDeriveBytes(
				passPhrase,
				saltValueBytes,
				hashAlgorithm,
				passwordIterations);

			// Use the password to generate pseudo-random bytes for the encryption
			// key. Specify the size of the key in bytes (instead of bits).
			key = password.GetBytes(keySize / 8);
		}

		private void YellowPepper_Load(object sender, System.EventArgs e)
		{
		
		}

		public string Encriptar(string original, string key)  
		{
			// Allgoritmo 3des
			TripleDESCryptoServiceProvider des;
			MD5CryptoServiceProvider hashmd5;
 
			byte[] keyhash, buff;  
			string encrypted;  
  
			hashmd5 = new MD5CryptoServiceProvider();  
			keyhash = hashmd5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(key));  

			hashmd5 = null;  
			des = new TripleDESCryptoServiceProvider();  
  
			des.Key = keyhash;  
			des.Mode = CipherMode.ECB;  
  
			buff = ASCIIEncoding.ASCII.GetBytes(original);  
			encrypted = Convert.ToBase64String(  
  
				des.CreateEncryptor().TransformFinalBlock(buff, 0, buff.Length)  
  
				);  
  
			return encrypted;  
		}  
  
		public string Desencriptar(string encrypted, string key)  
		{
			TripleDESCryptoServiceProvider des;  
			MD5CryptoServiceProvider hashmd5;  
			byte[] keyhash, buff;  
			string decrypted;  
			hashmd5 = new MD5CryptoServiceProvider();  
  
			keyhash = hashmd5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(key));  
			hashmd5 = null;  
			des = new TripleDESCryptoServiceProvider();  
			des.Key = keyhash;  
  
			des.Mode = CipherMode.ECB;  
			buff = Convert.FromBase64String(encrypted);  
			decrypted = ASCIIEncoding.ASCII.GetString(  
				des.CreateDecryptor().TransformFinalBlock(buff, 0, buff.Length));  
  
			return decrypted;
		}

		DataSet miDataSet = new DataSet();
		private void btCliente_Click(object sender, System.EventArgs e)
		{
			FileInfo fi = new FileInfo(txtPago.Text);
			if (fi.Name.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese archivo a subir", "Información");
				return;
			}

			string   passPhrase         = "Pas5pr@se";        // can be any string
			string   saltValue          = "s@1tValue";        // can be any string
			string   hashAlgorithm      = "SHA1";             // can be "MD5"
			int      passwordIterations = 3;                  // can be any number
			string   initVector         = "@1B2c3D4e5F6g7H8"; // must be 16 bytes
			int      keySize            = 192;                // can be 192 or 128

			string MiTexto = "";
			StreamReader sr = new StreamReader(txtPago.Text);
			MiTexto = sr.ReadToEnd();
			sr.Close();
			string stDecodificado = DecryptWith3DES(MiTexto, passPhrase, saltValue, hashAlgorithm, 
				passwordIterations, initVector, keySize);
			StreamWriter sw = new StreamWriter(fi.DirectoryName + "\\Desencriptar.csv");
			sw.Write(stDecodificado);
			sw.Flush();
			sw.Close();

			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Extended Properties=Text;" +
				"Data Source=" + fi.DirectoryName;
			OleDbDataAdapter myData = new OleDbDataAdapter(
				"SELECT * FROM " + "Desencriptar.csv", strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir el Archivo: " + txtPago.Text);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
			File.Delete(fi.DirectoryName + "\\Desencriptar.csv");
		}
		CultureInfo us = new CultureInfo("en-US");

		private void btSubir_Click(object sender, System.EventArgs e)
		{
			string stVerifica = ", 1";
			if (!chkVerifica.Checked) stVerifica = ", 0";
			else
			{
				string stBorrar = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorrar, true);
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
          if (dr.Cells[0].Value == DBNull.Value) continue;
					string stCodCliente = dr.Cells[0].Value.ToString().Trim();
					string stNumero = dr.Cells[1].Value.ToString().Trim();
					string stEstablecimiento = dr.Cells[2].Value.ToString().Trim();
					string stEmision = dr.Cells[3].Value.ToString().Trim();
					string stFecha = dr.Cells[4].Value.ToString().Trim();
					string stCodRecibe = dr.Cells[5].Value.ToString().Trim();
					string stCanal = dr.Cells[6].Value.ToString().Trim();
					string stOrigen = dr.Cells[7].Value.ToString().Trim();
					string stDestino = dr.Cells[8].Value.ToString().Trim();
					string stCarrier = dr.Cells[9].Value.ToString().Trim();
					double dSubTotal = double.Parse(dr.Cells[10].Value.ToString());
					double dIva = double.Parse(dr.Cells[11].Value.ToString());
					double dIce = double.Parse(dr.Cells[12].Value.ToString());
					double dTotal = double.Parse(dr.Cells[13].Value.ToString());
					string stBorrar = dr.Cells[14].Value.ToString().Trim();
					string stAutorizacion = dr.Cells[15].Value.ToString().Trim();
					string stFechaCaducidad = dr.Cells[16].Value.ToString().Trim();
					string stCentroCosto = dr.Cells[17].Value.ToString().Trim();

					if (stBorrar.ToUpper() == "ANULADO") stBorrar = "', 1";
					else stBorrar = "', 0";
					string stFechaHora = stFecha.Substring(0, 8) + ' ';
					if (stFecha.Length >= 11) stFechaHora += stFecha.Substring(9, 2)+":";
					else stFechaHora += "00:";
					if (stFecha.Length >= 13) stFechaHora += stFecha.Substring(11, 2)+":";
					else stFechaHora += "00:";
					if (stFecha.Length >= 15) stFechaHora += stFecha.Substring(13, 2);
					else stFechaHora += "00";
					stFecha = stFechaHora;
					string stSerie = stEstablecimiento.PadLeft(3, '0') + stEmision.PadLeft(3, '0');
					string stExec = "Exec YPFactura '" + stCodCliente
						+ "', '" + stNumero + "', '" + stSerie + "', '"
						+ stFecha + "', '" + stCodRecibe + "', '" + stCanal + " " + stOrigen 
						+ " " + stDestino + " ', " + dSubTotal.ToString("0.0000", us)
						+ ", " + dIva.ToString("0.0000", us) + ", " + dIce.ToString("0.0000", us) 
						+ ", " + dTotal.ToString("0.0000", us) + ", '" + stCarrier 
						+ stBorrar + ", '" + stAutorizacion + "', '" + stFechaCaducidad + "', '"
						+ stCentroCosto + "'" + stVerifica;

					Funcion.EjecutaSQL(cdsCliente, stExec, true);
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Información");
				}
			}
			if (chkVerifica.Checked) MuestraErrores();
			else MessageBox.Show("Notas de venta subidas", "Información");
		}
		private void MuestraErrores()
		{
			string stContar = "Select Count(*) From ErrorSalida";
			int iCuenta = Funcion.iEscalarSQL(cdsCliente, stContar, true);
			if (iCuenta == 0) 
			{
				MessageBox.Show("No hay errores en el archivo", "Información");
				return;
			}

			TesterError miError = new TesterError();
			miError.MdiParent = this.MdiParent;
			miError.Show();
		}

		private void btArchivo_Click(object sender, System.EventArgs e)
		{
			this.openFileDialog.Filter = "csv (*.csv)|*.csv";
			this.openFileDialog.InitialDirectory = @"c:\yellow\";
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				string stArchivo = this.openFileDialog.FileName.ToString();
				txtPago.Text = stArchivo;
			}
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btDesencriptar_Click(object sender, System.EventArgs e)
		{
			string   plainText          = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";    // original plaintext
      
			string   passPhrase         = "Pas5pr@se";        // can be any string
			string   saltValue          = "s@1tValue";        // can be any string
			string   hashAlgorithm      = "SHA1";             // can be "MD5"
			int      passwordIterations = 3;                  // can be any number
			string   initVector         = "@1B2c3D4e5F6g7H8"; // must be 16 bytes
			int      keySize            = 192;                // can be 192 or 128
      
			MessageBox.Show(plainText);
			string  cipherText = Encrypt(plainText,
				passPhrase,
				saltValue,
				hashAlgorithm,
				passwordIterations,
				initVector,
				keySize);

			MessageBox.Show(cipherText);
        
			plainText          = Decrypt(cipherText,
				passPhrase,
				saltValue,
				hashAlgorithm,
				passwordIterations,
				initVector,
				keySize);

			MessageBox.Show(plainText);
		}

		private void DecryptFile(string inFile)
		{
			// Create instance of Rijndael for
			// symetric decryption of the data.
			RijndaelManaged rjndl = new RijndaelManaged();
			rjndl.KeySize = 192;
			rjndl.BlockSize = 192;
			rjndl.Mode = CipherMode.CBC;
			rjndl.Padding = PaddingMode.None;

			// Create byte arrays to get the length of
			// the encrypted key and IV.
			// These values were stored as 4 bytes each
			// at the beginning of the encrypted package.
			byte[] LenK = new byte[4];
			byte[] LenIV = new byte[4];

			// Consruct the file name for the decrypted file.
			string outFile = inFile.Substring(0, inFile.LastIndexOf(".")) + ".txt";

			// Use FileStream objects to read the encrypted
			// file (inFs) and save the decrypted file (outFs).
			using (FileStream inFs = new FileStream(inFile, FileMode.Open))
			{

				inFs.Seek(0, SeekOrigin.Begin);
				inFs.Seek(0, SeekOrigin.Begin);
				inFs.Read(LenK, 0, 3);
				inFs.Seek(4, SeekOrigin.Begin);
				inFs.Read(LenIV, 0, 3);

				// Convert the lengths to integer values.
				int lenK = BitConverter.ToInt32(LenK, 0);
				int lenIV = BitConverter.ToInt32(LenIV, 0);

				// Determine the start postition of
				// the ciphter text (startC)
				// and its length(lenC).
				int startC = lenK + lenIV + 8;
				int lenC = (int)inFs.Length - startC;

				// Create the byte arrays for
				// the encrypted Rijndael key,
				// the IV, and the cipher text.
				byte[] KeyEncrypted = new byte[lenK];
				byte[] IV = new byte[lenIV];

				// Extract the key and IV
				// starting from index 8
				// after the length values.
				inFs.Seek(8, SeekOrigin.Begin);
				inFs.Read(KeyEncrypted, 0, lenK);
				inFs.Seek(8 + lenK, SeekOrigin.Begin);
				inFs.Read(IV, 0, lenIV);

				// Use RSACryptoServiceProvider
				// to decrypt the Rijndael key.
				byte[] KeyDecrypted = null;//rsa.Decrypt(KeyEncrypted, false);

				// Decrypt the key.
				ICryptoTransform transform = rjndl.CreateDecryptor(KeyDecrypted, IV);

				// Decrypt the cipher text from
				// from the FileSteam of the encrypted
				// file (inFs) into the FileStream
				// for the decrypted file (outFs).
				using (FileStream outFs = new FileStream(outFile, FileMode.Create))
				{

					int count = 0;
					int offset = 0;

					// blockSizeBytes can be any arbitrary size.
					int blockSizeBytes = rjndl.BlockSize / 8;
					byte[] data = new byte[blockSizeBytes];


					// By decrypting a chunk a time,
					// you can save memory and
					// accommodate large files.

					// Start at the beginning
					// of the cipher text.
					inFs.Seek(startC, SeekOrigin.Begin);
					using (CryptoStream outStreamDecrypted = new CryptoStream(outFs, transform, CryptoStreamMode.Write))
					{
						do
						{
							count = inFs.Read(data, 0, blockSizeBytes);
							offset += count;
							outStreamDecrypted.Write(data, 0, count);

						}
						while (count > 0);

						outStreamDecrypted.FlushFinalBlock();
						outStreamDecrypted.Close();
					}
					outFs.Close();
				}
				inFs.Close();
			}
		}

		private void btDesencriptar_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			DecryptFile(txtPago.Text);
		}

		private void btSubirArchivo_Click(object sender, System.EventArgs e)
		{
			string stVerifica = ", 1";
			if (!chkVerifica.Checked) stVerifica = ", 0";
			else
			{
				string stBorrar = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorrar, true);
			}

			StreamReader objReader = new StreamReader(txtPago.Text);
			string stLinea = "";
			stLinea = objReader.ReadLine();
			while (stLinea != null)
			{
				try
				{
					stLinea = objReader.ReadLine();
					if (stLinea == null) break;
					string [] stCodigos = stLinea.Split(',');
					string stCodCliente = stCodigos[0];
					string stNumero = stCodigos[1];
					string stEstablecimiento = stCodigos[2];
					string stEmision = stCodigos[3];
					string stFecha = stCodigos[4];
					string stCodRecibe = stCodigos[5];
					string stCanal = stCodigos[6];
					string stOrigen = stCodigos[7];
					string stDestino = stCodigos[8];
					string stCarrier = stCodigos[9];
					double dSubTotal = double.Parse(stCodigos[10]);
					double dIva = double.Parse(stCodigos[11]);
					double dIce = double.Parse(stCodigos[12]);
					double dTotal = double.Parse(stCodigos[13]);
					string stBorrar = stCodigos[14];
					string stAutorizacion = stCodigos[15];
					string stFechaCaducidad = stCodigos[16];
					string stCentroCosto = stCodigos[17];

					if (stBorrar.ToUpper() == "ANULADO") stBorrar = "', 1";
					else stBorrar = "', 0";
					string stFechaHora = stFecha.Substring(0, 8) + ' ';
					if (stFecha.Length >= 11) stFechaHora += stFecha.Substring(9, 2)+":";
					else stFechaHora += "00:";
					if (stFecha.Length >= 13) stFechaHora += stFecha.Substring(11, 2)+":";
					else stFechaHora += "00:";
					if (stFecha.Length >= 15) stFechaHora += stFecha.Substring(13, 2);
					else stFechaHora += "00";
					stFecha = stFechaHora;
					string stSerie = stEstablecimiento.PadLeft(3, '0') + stEmision.PadLeft(3, '0');
					string stExec = "Exec YPFactura '" + stCodCliente
						+ "', '" + stNumero + "', '" + stSerie + "', '"
						+ stFecha + "', '" + stCodRecibe + "', '" + stCanal + " " + stOrigen 
						+ " " + stDestino + " ', " + dSubTotal.ToString("0.0000", us)
						+ ", " + dIva.ToString("0.0000", us) + ", " + dIce.ToString("0.0000", us) 
						+ ", " + dTotal.ToString("0.0000", us) + ", '" + stCarrier 
						+ stBorrar + ", '" + stAutorizacion + "', '" + stFechaCaducidad + "', '"
						+ stCentroCosto + "'" + stVerifica;

					Funcion.EjecutaSQL(cdsCliente, stExec, true);
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Información");
				}
			}
			objReader.ReadLine();
			objReader.Close();
			if (chkVerifica.Checked) MuestraErrores();
			else MessageBox.Show("Notas de venta subidas", "Información");
		}

		private void txtPago_ValueChanged(object sender, System.EventArgs e)
		{
      
		}
	}
}
