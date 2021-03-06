﻿using System.Collections.Generic;
using System.Data;
using ElecDocServices.Helpers;

namespace ElecDocServices.Providers
{
    internal interface IFacElecInterface
    {
        DataTable DocHeader { get; set; }
        DataTable DocDetail { get; set; }
        DataTable DocProvider { get; set; }

        /// <summary>
        /// 
        /// </summary>
        List<Parameter> RegistrarDocumento();

        /// <summary>
        /// 
        /// </summary>
        List<Parameter> RegistrarDocNC();

        /// <summary>
        /// 
        /// </summary>
        List<Parameter> AnularDocumento();

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        List<Parameter> AnularDocNC();

        /// <summary>
        /// 
        /// </summary>
        List<Parameter> ObtenerDocumento();
    }
}
