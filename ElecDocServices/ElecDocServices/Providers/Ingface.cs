using System;
using System.Collections.Generic;
using System.Data;
using ElecDocServices.Helpers;
using ElecDocServices.IngfaceServices;

namespace ElecDocServices.Providers
{
    internal class Ingface : IFacElecInterface
    {
        private Utils utl = new Utils();

        public DataTable DocHeader { get; set; }
        public DataTable DocDetail { get; set; }
        public DataTable DocProvider { get; set; }

        public List<Parameter> RegistrarDocumento()
        {
            responseDte res = new responseDte();
            requestDte reg = new requestDte();

            try
            {
                ingface service = new ingface();
                reg = ConstruirDatos();
                res = service.registrarDte(reg);
            }
            catch (Exception ex)
            {
                res.valido = false;
                res.descripcion = ex.Message;
            }

            return ObtenerResultado(res, reg, "registrarDte");
        }

        public List<Parameter> ObtenerDocumento()
        {
            throw new NotImplementedException();
        }

        public List<Parameter> AnularDocumento()
        {
            throw new NotImplementedException();
        }


        private requestDte ConstruirDatos()
        {
            requestDte datos = new requestDte();
            dte d = new dte();

            if (this.DocHeader.Rows.Count > 0)
            {
                DataRow r = this.DocHeader.Rows[0];

                //Configuracion Inicial
                d.codigoEstablecimiento = utl.convertirString(r["Empresa"]);
                d.idDispositivo = utl.convertirString(r["Caja"]);
                d.serieAutorizada = utl.convertirString(r["Serie"]);
                d.numeroResolucion = utl.convertirString(r["Resolucion"]);
                d.fechaResolucion = utl.convertirDateTime(this.DocProvider.Rows[0]["Fecha"]);
                d.fechaResolucionSpecified = false;
                d.tipoDocumento = utl.convertirString(r["Documento"]);
                d.serieDocumento = utl.convertirString(r["Serie"]);
                //d.nitGFACE = utl.convertirString(this.DocProvider.Rows[0][""]);

                //Datos de documento y cliente
                d.numeroDocumento = utl.convertirString(r["DocNo"]);
                d.fechaDocumento = utl.convertirDateTime(r["Fecha"]);
                d.fechaDocumentoSpecified = true;
                d.estadoDocumento = "ACTIVO";
                d.codigoMoneda = utl.convertirString(r["Divisa"]);
                d.tipoCambio = utl.convertirDouble(r["TasaCambio"]);
                d.tipoCambioSpecified = true;
                d.nitComprador = utl.convertirString(r["NITCliente"]);
                d.nombreComercialComprador = utl.convertirString(r["NombreCliente"]);
                d.direccionComercialComprador = utl.convertirString(r["DireccionCliente"]);
                d.telefonoComprador = "N/A";
                d.correoComprador = "N/A";
                d.municipioComprador = "N/A";
                d.departamentoComprador = "N/A";

                //Datos de valores de documento
                d.importeBruto = utl.convertirDouble(r["ValorNeto"]);
                d.importeBrutoSpecified = true;
                d.detalleImpuestosIva = utl.convertirDouble(r["IVA"]);
                d.detalleImpuestosIvaSpecified = true;
                d.importeNetoGravado = utl.convertirDouble(r["Total"]);
                d.importeTotalExento = utl.convertirDouble(r["Exento"]);
                d.montoTotalOperacion = utl.convertirDouble(r["Total"]);

                //Datos de vendedor
                d.nitVendedor = "N/A";
                d.nombreComercialRazonSocialVendedor = "N/A";
                d.direccionComercialVendedor = "N/A";
                d.municipioVendedor = "N/A";
                d.departamentoVendedor = "N/A";
                d.observaciones = "N/A";
                d.regimenISR = "N/A";

                //Campos Personalizados

                //Detalle de documento
                if (this.DocDetail.Rows.Count > 0)
                {
                    List<detalleDte> deta = new List<detalleDte>();

                    foreach (DataRow rf in DocDetail.Rows)
                    {
                        detalleDte dt = new detalleDte();

                        dt.cantidad = utl.convertirDouble(rf["Cantidad"]);
                        dt.cantidadSpecified = true;
                        dt.codigoProducto = utl.convertirString(rf["Codigo"]);
                        dt.descripcionProducto = utl.convertirString(rf["Descripcion"]);
                        dt.precioUnitario = utl.convertirDouble(rf["PrecioUnitario"]);
                        dt.precioUnitarioSpecified = true;
                        dt.montoBruto = utl.convertirDouble(rf["Importe"]);
                        dt.montoBrutoSpecified = true;
                        dt.detalleImpuestosIva = utl.convertirDouble(rf["IVA"]);
                        dt.detalleImpuestosIvaSpecified = true;
                        dt.importeNetoGravado = utl.convertirDouble(rf["Total"]);
                        dt.importeNetoGravadoSpecified = true;
                        dt.importeTotalOperacion = utl.convertirDouble(rf["Total"]);
                        dt.importeTotalOperacionSpecified = true;
                        dt.unidadMedida = utl.convertirString(rf["Metrica"]);
                        dt.tipoProducto = utl.convertirString(rf["Tipo"]);

                        //Campos personalizados

                        deta.Add(dt);
                    }

                    //Adicion a Encabezado
                    d.detalleDte = deta.ToArray();
                }

                datos.dte = d;
                datos.usuario = utl.convertirString(this.DocProvider.Rows[0]["Auth"]);
                datos.clave = utl.convertirString(this.DocProvider.Rows[0]["Auth"]);

            }

            return new requestDte();
        }


        private List<Parameter> ObtenerResultado(object res, object request, string function)
        {
            return new List<Parameter>();
        }
    }
}
