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
            requestDte re = new requestDte();

            return re;
        }


        private List<Parameter> ObtenerResultado(object res, object request, string function)
        {
            return new List<Parameter>();
        }
    }
}
