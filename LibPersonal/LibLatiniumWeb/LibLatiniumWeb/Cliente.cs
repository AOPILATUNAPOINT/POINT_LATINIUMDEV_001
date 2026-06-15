using System;
using DevExpress.Xpo;
using DevExpress.Xpo.DB;
namespace INFOELECT2007
{
  public class Cliente : XPLiteObject
  {
    int fidCliente;
    [Key(true)]
    public int idCliente
    {
      get { return fidCliente; }
      set { SetPropertyValue<int>("idCliente", ref fidCliente, value); }
    }
    XPCollection coleccion = new XPCollection(typeof(Cliente));
    
    int fidGrupoCliente;
    public int idGrupoCliente
    {
      get { return fidGrupoCliente; }
      set { SetPropertyValue<int>("idGrupoCliente", ref fidGrupoCliente, value); }
    }
    int fidTipoPrecio;
    public int idTipoPrecio
    {
      get { return fidTipoPrecio; }
      set { SetPropertyValue<int>("idTipoPrecio", ref fidTipoPrecio, value); }
    }
    string fNombre;
    [Size(50)]
    public string Nombre
    {
      get { return fNombre; }
      set { SetPropertyValue<string>("Nombre", ref fNombre, value); }
    }
    string fCodigo;
    [Size(15)]
    public string Codigo
    {
      get { return fCodigo; }
      set { SetPropertyValue<string>("Codigo", ref fCodigo, value); }
    }
    bool fAmbos;
    public bool Ambos
    {
      get { return fAmbos; }
      set { SetPropertyValue<bool>("Ambos", ref fAmbos, value); }
    }
    bool fProveedor;
    public bool Proveedor
    {
      get { return fProveedor; }
      set { SetPropertyValue<bool>("Proveedor", ref fProveedor, value); }
    }
    public Cliente(Session session) : base(session) { }
    public Cliente() : base(Session.DefaultSession) { }
    public override void AfterConstruction() { base.AfterConstruction(); }
  }

}
