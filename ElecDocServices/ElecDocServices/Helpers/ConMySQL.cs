using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using MySql.Data.MySqlClient;
using System.IO;

namespace ElecDocServices.Helpers
{
    /******************************************************
    * En esta clase se definen todos lo metodos para  
    * realizar la conexión o conexiones a las base de 
    * datos y su desconexión
    ******************************************************/

    /// <summary>
    /// Clase que contiene las funciones para operaciones de consulta con base de datos MySQL
    /// </summary>
    internal class ConMySQL
    {
        //Clases Globales
        private Utils utl = new Utils();
        private Errors err = null;

        //Variables Globales
        private MySqlConnection sqlConn = null;
        private MySqlDataAdapter daAdapter = null;
        private MySqlCommand sqlcomm = null;

        //Variables Globales
        private string ConfigPath = null;
        private string SystemUser = null;


        private string sql = null;
        private DataTable dtTable = null;



        /// <summary>
        /// Inicializa una instancia de ConMySQL con target esquema del Dealer del sistema
        /// </summary>
        public ConMySQL(string Config, string SysUser)
        {
            if (!File.Exists(Config))
            {
                throw new Exception("Config file does not exists");
            }

            this.ConfigPath = Config;
            this.SystemUser = SysUser;

            string path = utl.convertirString(utl.getConfigValue(this.ConfigPath, "SQL", "errors"));
            err = new Errors(this.SystemUser, "SQL", path);
        }


        ///<summary>
        ///Obtiene el target de Conexion a la base SAL /GT
        ///</summary>
        //Conexion a la base SAL /GT
        private string obtenerConnectionString()
        {
            string user = utl.convertirString(utl.getConfigValue(this.ConfigPath, "USER", "mysql"));
            string pass = utl.convertirString(utl.getConfigValue(this.ConfigPath, "PASSWORD", "mysql"));
            string server = utl.convertirString(utl.getConfigValue(this.ConfigPath, "SERVER", "mysql"));
            string database = utl.convertirString(utl.getConfigValue(this.ConfigPath, "DATABASE", "mysql"));

            string constring = "Data Source=" + server + ";Database=" + database + ";User Id=" + user + ";Password=" + pass + ";";

            return constring;
        }

        
        ///<summary>
        ///Apertura de Conexion
        ///</summary>
        //Apertura de Conexion
        private void conectar()
        {
            sqlConn = new MySqlConnection(obtenerConnectionString());

            if (sqlConn != null && sqlConn.State != ConnectionState.Open)
            {
                sqlConn.Open();
            }
        }


        ///<summary>
        ///Desconexion de la base de datos.
        ///</summary>
        //Desconexion de la base de datos
        private void desconectar()
        {
            if (sqlConn != null && sqlConn.State == ConnectionState.Open)
            {
                sqlConn.Close();
            }   
        }




        #region consultas



        ///<summary>
        ///Obtiene Datos por consulta
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="campos">Campos a traer en el source de datos</param>
        ///<param name="where">Sentencia SQL WHERE de la data para filtrado (Syntax SQL) (Opcional)</param>
        ///<param name="groupby">Sentencia SQL GROUP BY de la data (Syntax SQL) (Opcional)</param>
        ///<param name="ordby">Sentencia SQL ORDER BY de la data (campos) (Syntax SQL) (Opcional)</param>
        //Obtiene Datos por consulta
        public DataTable obtenerDatos(string tabla, string campos, string where = "", string groupby = "", string ordby = "", int? limit = null)
        {
            try
            {
                this.conectar();
                sql = "select " + campos + " from " + tabla + " ";
                sql += (where != "") ? " where " + where : "";
                sql += (groupby != "") ? " group by " + groupby : "";
                sql += (ordby != "") ? " order by " + ordby : "";
                sql += (limit != null) ? " limit " + utl.convertirString(limit) : "";
                sql += " ;";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                this.desconectar();
                return dtTable;
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Select Datos - MySQL");
                this.desconectar();
                return new DataTable();
            }
        }


        ///<summary>
        ///Obtiene Datos por consulta con join
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="campos">Campos a traer en el source de datos</param>
        ///<param name="join">Arreglo de cadenas para Sentencias SQL join (Syntax SQL)</param>
        ///<param name="where">Sentencia SQL WHERE de la data para filtrado (Syntax SQL) (Opcional)</param>
        ///<param name="groupby">Sentencia SQL GROUP BY de la data (Syntax SQL) (Opcional)</param>
        ///<param name="ordby">Sentencia SQL ORDER BY de la data (campos) (Syntax SQL) (Opcional)</param>
        //Obtiene Datos por consulta con join
        public DataTable obtenerDatos(string tabla, string campos, string[] join, string where = "", string groupby = "", string ordby = "", int? limit = null)
        {
            try
            {
                this.conectar();
                sql = "select " + campos + " from " + tabla + " ";

                for (int i = 0; i < join.Length; i++)
                {
                    sql += " " + join[i] + " ";
                }

                sql += (where != "") ? " where " + where : "";
                sql += (groupby != "") ? " group by " + groupby : "";
                sql += (ordby != "") ? " order by " + ordby : "";
                sql += (limit != null) ? " limit " + utl.convertirString(limit) : "";
                sql += " ;";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                this.desconectar();
                return dtTable;
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Select Datos - MySQL");
                this.desconectar();
                return new DataTable();
            }
        }


        ///<summary>
        ///Obtener Datos por procedimiento
        ///</summary>
        ///<param name="procedimiento">Nombre de Procedimiento Almacenado</param>
        ///<param name="datos">Arreglo de datos parametros del procedimiento almacenado</param>
        //Obtener Datos por procedimiento
        public DataTable obtenerDatosSp(string procedimiento, string[] datos)
        {
            try
            {
                this.conectar();
                sql = "CALL " + procedimiento + " ( ";
                for (int i = 0; i < datos.Count(); i++)
                {
                    if (i != 0) { sql += ", "; }
                    sql += (datos[i] != "") ? "'" + datos[i] + "'" : "''";
                }
                sql += " ) ;";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);

                //Obtencion de Warnings de Query
                daAdapter = new MySqlDataAdapter("show warnings;", sqlConn);
                DataTable dtw = new DataTable();
                daAdapter.Fill(dtw);

                if (dtw.Rows.Count > 0)
                {
                    foreach (DataRow r in dtw.Rows)
                    {
                        err.AddErrors("Warning: " + utl.convertirString(r["Message"]), "Cod. Error:" + utl.convertirString(r["Code"]), "Query Procedure - MySQL", sql, this.SystemUser);
                    }
                }

                this.desconectar();

                return dtTable;
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Select Datos SP - MySQL");
                this.desconectar();
                return new DataTable();
            }
        }


        ///<summary>
        ///Obtener Datos (Una Fila) por procedimiento 
        ///</summary>
        ///<param name="procedimiento">Nombre de Procedimiento Almacenado</param>
        ///<param name="datos">Arreglo de datos parametros del procedimiento almacenado</param>
        //Obtener Datos (Una Fila) por procedimiento 
        public DataRow obtenerUnDatoSp(string procedimiento, string[] datos)
        {
            try
            {
                this.conectar();
                sql = "CALL " + procedimiento + " ( ";
                for (int i = 0; i < datos.Count(); i++)
                {
                    if (i != 0) { sql += ", "; }
                    sql += (datos[i] != "") ? "'" + datos[i] + "'" : "''";
                }
                sql += " ) ;";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                this.desconectar();
                return dtTable.Rows[0];
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Select Datos SP - MySQL");
                this.desconectar();
                return null;
            }
        }


        ///<summary>
        ///Obtiene un dato especifico por consulta
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="campo">Campo a traer en el source de datos</param>
        ///<param name="where">Sentencia SQL WHERE de la data para filtrado (Syntax SQL) (Opcional)</param>
        ///<param name="groupby">Sentencia SQL GROUP BY de la data (Syntax SQL) (Opcional)</param>
        ///<param name="ordby">Sentencia SQL ORDER BY de la data (campos) (Syntax SQL) (Opcional)</param>
        //Obtiene dato especifico por consulta
        public object obtenerUnDato(string tabla, string campo, string where = "", string groupby = "", string ordby = "")
        {
            object dato = null;

            try
            {
                this.conectar();
                sql = "select " + campo + " from " + tabla + " ";
                sql += (where != "") ? " where " + where : "";
                sql += (groupby != "") ? " group by " + groupby : "";
                sql += (ordby != "") ? " order by " + ordby : "";
                sql += " ;";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                this.desconectar();

                if (dtTable.Rows.Count > 0)
                {
                    dato = dtTable.Rows[0][0];
                }

                return dato;
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Select Datos - MySQL");
                this.desconectar();
                return null;
            }
        }


        ///<summary>
        ///Obtiene un dato especifico por consulta con join
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="campo">Campo a traer en el source de datos</param>
        ///<param name="join">Arreglo de cadenas para Sentencias SQL join (Syntax SQL)</param>
        ///<param name="where">Sentencia SQL WHERE de la data para filtrado (Syntax SQL) (Opcional)</param>
        ///<param name="groupby">Sentencia SQL GROUP BY de la data (Syntax SQL) (Opcional)</param>
        ///<param name="ordby">Sentencia SQL ORDER BY de la data (campos) (Syntax SQL) (Opcional)</param>
        //Obtiene dato especifico por consulta con join
        public object obtenerUnDato(string tabla, string campo, string[] join, string where = "", string groupby = "", string ordby = "")
        {
            object dato = null;

            try
            {
                this.conectar();
                sql = "select " + campo + " from " + tabla + " ";

                for (int i = 0; i < join.Length; i++)
                {
                    sql += " " + join[i] + " ";
                }

                sql += (where != "") ? " where " + where : "";
                sql += (groupby != "") ? " group by " + groupby : "";
                sql += (ordby != "") ? " order by " + ordby : "";
                sql += " ;";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                this.desconectar();
                
                if (dtTable.Rows.Count > 0)
                {
                    dato = dtTable.Rows[0][0];
                }

                return dato;
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Select Datos - MySQL");
                this.desconectar();
                return null;
            }
        }


        ///<summary>
        ///Obtener Conteo de Consulta
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="where">Sentencia SQL WHERE de la data para filtrado (Syntax SQL) (Opcional)</param>
        ///<param name="groupby">Sentencia SQL GROUP BY de la data (Syntax SQL) (Opcional)</param>
        ///<param name="ordby">Sentencia SQL ORDER BY de la data (campos) (Syntax SQL) (Opcional)</param>
        //Obtener Conteo de Consulta
        public int obtenerConteo(string tabla, string where = "", string groupby = "", string ordby = "")
        {
            int conteo = 0;
            try
            {
                this.conectar();
                sql = "select count(*) as conteo from " + tabla + " ";
                sql += (where != "") ? " where " + where : "";
                sql += (groupby != "") ? " group by " + groupby : "";
                sql += (ordby != "") ? " order by " + ordby : "";
                sql += " ;";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                this.desconectar();
                if (dtTable.Rows.Count > 0) { conteo = utl.convertirInt(dtTable.Rows[0]["conteo"]); }
                return conteo;
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Select Count Datos - MySQL");
                this.desconectar();
                return conteo;
            }
        }


        ///<summary>
        ///Obtener Conteo de Consulta con join
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="join">Arreglo de cadenas para Sentencias SQL join (Syntax SQL)</param>
        ///<param name="where">Sentencia SQL WHERE de la data para filtrado (Syntax SQL) (Opcional)</param>
        ///<param name="groupby">Sentencia SQL GROUP BY de la data (Syntax SQL) (Opcional)</param>
        ///<param name="ordby">Sentencia SQL ORDER BY de la data (campos) (Syntax SQL) (Opcional)</param>
        //Obtener Conteo de Consulta con join
        public int obtenerConteo(string tabla, string[] join, string where = "", string groupby = "", string ordby = "")
        {
            int conteo = 0;
            try
            {
                this.conectar();
                sql = "select count(*) as conteo from " + tabla + " ";

                for (int i = 0; i < join.Length; i++)
                {
                    sql += " " + join[i] + " ";
                }

                sql += (where != "") ? " where " + where : "";
                sql += (groupby != "") ? " group by " + groupby : "";
                sql += (ordby != "") ? " order by " + ordby : "";
                sql += " ;";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                this.desconectar();
                if (dtTable.Rows.Count > 0) { conteo = utl.convertirInt(dtTable.Rows[0]["conteo"]); }
                return conteo;
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Select Datos - MySQL");
                this.desconectar();
                return conteo;
            }
        }


        ///<summary>
        ///Obtener Max de Consulta
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="campo">Campo a evaluar el MAX en el source de datos</param>
        ///<param name="where">Sentencia SQL WHERE de la data para filtrado (Syntax SQL) (Opcional)</param>
        ///<param name="groupby">Sentencia SQL GROUP BY de la data (Syntax SQL) (Opcional)</param>
        ///<param name="ordby">Sentencia SQL ORDER BY de la data (campos) (Syntax SQL) (Opcional)</param>
        //Obtener Max de Consulta
        public int? obtenerMax(string tabla, string campo, string where = "", string groupby = "", string ordby = "")
        {
            int max = 0;
            try
            {
                this.conectar();
                sql = "select MAX(" + campo + ") as max from " + tabla + " ";
                sql += (where != "") ? " where " + where : "";
                sql += (groupby != "") ? " group by " + groupby : "";
                sql += (ordby != "") ? " order by " + ordby : "";
                sql += " ;";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                this.desconectar();
                if (dtTable.Rows.Count > 0) { max = utl.convertirInt(dtTable.Rows[0]["max"]); }
                return max;
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Select Max Datos - MySQL");
                this.desconectar();
                return null;
            }
        }


        ///<summary>
        ///Obtener Max de Consulta con join
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="campo">Campo a evaluar el MAX en el source de datos</param>
        ///<param name="join">Arreglo de cadenas para Sentencias SQL join (Syntax SQL)</param>
        ///<param name="where">Sentencia SQL WHERE de la data para filtrado (Syntax SQL) (Opcional)</param>
        ///<param name="groupby">Sentencia SQL GROUP BY de la data (Syntax SQL) (Opcional)</param>
        ///<param name="ordby">Sentencia SQL ORDER BY de la data (campos) (Syntax SQL) (Opcional)</param>
        //Obtener Max de Consulta con join
        public int? obtenerMax(string tabla, string campo, string[] join, string where = "", string groupby = "", string ordby = "")
        {
            int max = 0;
            try
            {
                this.conectar();
                sql = "select MAX(" + campo + ") as max from " + tabla + " ";

                for (int i = 0; i < join.Length; i++)
                {
                    sql += " " + join[i] + " ";
                }

                sql += (where != "") ? " where " + where : "";
                sql += (groupby != "") ? " group by " + groupby : "";
                sql += (ordby != "") ? " order by " + ordby : "";
                sql += " ;";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                this.desconectar();
                if (dtTable.Rows.Count > 0) { max = utl.convertirInt(dtTable.Rows[0]["max"]); }
                return max;
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Select Max Datos - MySQL");
                this.desconectar();
                return null;
            }
        }



        #endregion



        #region procedimientos



        ///<summary>
        ///Ejecucion de procedimiento insert
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="campos">Arreglo con los nombres de los campos de la tabla</param>
        ///<param name="datos">Arreglo con los valores de los campos en el orden de los mismos</param>
        ///<param name="nonulo">Aceptan valores Nulos (Opcional) (Default: false)</param>
        //Ejecucion de procedimiento insert
        public bool execInsert(string tabla, string[] campos, object[] datos, bool nonulo = false)
        {
            bool retorno = false;

            try
            {
                this.conectar();
                sql = "insert into " + tabla + " ( ";
                for (int i = 0; i < campos.Length; i++)
                {
                    if (i != 0) { sql += ", "; }
                    sql += campos[i];
                }
                sql += " )  values  ( ";
                for (int i = 0; i < datos.Length; i++)
                {
                    if (i != 0) { sql += ", "; }
                    sql += utl.formatoDatosSQL(datos[i], nonulo);
                }
                sql += " ) ";
                sqlcomm = sqlConn.CreateCommand();
                sqlcomm.CommandText = sql;
                int a = sqlcomm.ExecuteNonQuery();
                retorno = (a > 0) ? true : false;

                this.desconectar();
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Insert - MySQL");
                retorno = false;
                this.desconectar();
            }

            return retorno;
        }


        ///<summary>
        ///Ejecucion de procedimiento insert
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="campos">Arreglo con los nombres de los campos de la tabla</param>
        ///<param name="datos">Lista de Arreglos con los valores de los campos en el orden de los mismos (Insert multiple)</param>
        ///<param name="nonulo">Aceptan valores Nulos (Opcional) (Default: false)</param>
        //Ejecucion de procedimiento insert
        public bool execInsert(string tabla, string[] campos, List<object[]> datos, bool nonulo = false)
        {
            bool retorno = false;

            try
            {
                this.conectar();
                sql = "insert into " + tabla + " ( ";
                for (int i = 0; i < campos.Length; i++)
                {
                    if (i != 0) { sql += ", "; }
                    sql += campos[i];
                }
                sql += " )  values  ";
                for (int d = 0; d < datos.Count; d++)
                {
                    object[] items = datos[d];

                    if (d != 0) { sql += ", "; }

                    sql += "( ";
                    for (int i = 0; i < items.Length; i++)
                    {
                        if (i != 0) { sql += ", "; }
                        sql += utl.formatoDatosSQL(items[i], nonulo);
                    }
                    sql += " ) ";
                }
                sqlcomm = sqlConn.CreateCommand();
                sqlcomm.CommandText = sql;
                int a = sqlcomm.ExecuteNonQuery();
                retorno = (a > 0) ? true : false;

                this.desconectar();
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Insert - MySQL");
                retorno = false;
                this.desconectar();
            }

            return retorno;
        }


        ///<summary>
        ///Ejecucion de procedimiento update
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="campos">Arreglo con los nombres de los campos de la tabla</param>
        ///<param name="datos">Arreglo con los valores de los campos en el orden de los mismos</param>
        ///<param name="where">Sentencia SQL WHERE de la data para filtrado (Syntax SQL)</param>
        ///<param name="nonulo">Aceptan valores Nulos (Opcional) (Default: false)</param>
        //Ejecucion de procedimiento update
        public bool execUpdate(string tabla, string[] campos, object[] datos, string where, bool nonulo = false)
        {
            bool retorno = false;

            try
            {
                this.conectar();

                sql = "update " + tabla + " set ";
                for (int i = 0; i < campos.Length; i++)
                {
                    if (i != 0) { sql += ", "; }

                    sql += campos[i] + " = " + utl.formatoDatosSQL(datos[i], nonulo);
                }
                sql += " where " + where + " ; ";

                sqlcomm = sqlConn.CreateCommand();
                sqlcomm.CommandText = sql;
                int a = sqlcomm.ExecuteNonQuery();
                retorno = (a > 0) ? true : false;

                this.desconectar();
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Update - MySQL");
                retorno = false;
                this.desconectar();
            }

            return retorno;
        }


        /// <summary>
        /// Ejecucion de procedimiento update
        /// </summary>
        /// <param name="tabla">Nombre de tabla de base de datos</param>
        /// <param name="campos">Nombre del campo de la tabla</param>
        /// <param name="datos">Valor del campo de la tabla</param>
        /// <param name="where">Sentencia SQL WHERE de la data para filtrado (Syntax SQL)</param>
        /// <param name="nonulo">Aceptan valores Nulos (Opcional) (Default: false)</param>
        /// <returns></returns>
        //Ejecucion de procedimiento update
        public bool execUpdate(string tabla, string campo, object dato, string where, bool nonulo = false)
        {
            bool retorno = false;

            try
            {
                this.conectar();

                sql = "update " + tabla + " set ";
                sql += campo + " = " + utl.formatoDatosSQL(dato, nonulo);
                sql += " where " + where + " ; ";

                sqlcomm = sqlConn.CreateCommand();
                sqlcomm.CommandText = sql;
                int a = sqlcomm.ExecuteNonQuery();
                retorno = (a > 0) ? true : false;

                this.desconectar();
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Update - MySQL");
                retorno = false;
                this.desconectar();
            }

            return retorno;
        }


        ///<summary>
        ///Ejecucion de procedimiento delete
        ///</summary>
        ///<param name="tabla">Nombre de tabla de base de datos</param>
        ///<param name="where">Sentencia SQL WHERE de la data para filtrado (Syntax SQL) (Opcional)</param>
        //Ejecucion de procedimiento delete
        public bool execDelete(string tabla, string where)
        {
            bool retorno = false;

            try
            {
                int r = obtenerDatos(tabla, "*", where).Rows.Count;

                if (r > 0)
                {
                    this.conectar();

                    sql = "delete from " + tabla + " ";
                    sql += " where " + where + " ; ";

                    sqlcomm = sqlConn.CreateCommand();
                    sqlcomm.CommandText = sql;
                    int a = sqlcomm.ExecuteNonQuery();
                    retorno = (a > 0) ? true : false;
                    this.desconectar();
                }
                else
                {
                    retorno = true;
                }
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Delete - MySQL");
                retorno = false;
                this.desconectar();
            }

            return retorno;
        }


        ///<summary>
        ///Ejecucion de funcion con retorno booleano
        ///</summary>
        ///<param name="fun">Nombre de la función</param>
        ///<param name="datos">Arreglo con los valores de la funcion</param>
        //Ejecucion de funcion con retorno booleano (1/0)
        public bool execBoolFunction(string fun, string[] datos)
        {
            bool retorno = false;

            try
            {
                this.conectar();
                sql = "select " + fun + " ( ";
                for (int i = 0; i < datos.Length; i++)
                {
                    if (i != 0) { sql += ", "; }
                    sql += "'" + datos[i] + "'";
                }
                sql += " ) as res ; ";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                int r = utl.convertirInt(dtTable.Rows[0]["res"]);
                retorno = (r == 1) ? true : false;

                this.desconectar();
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Boolean Function - MySQL");
                retorno = false;
                this.desconectar();
            }

            return retorno;
        }


        ///<summary>
        ///Ejecucion de funcion con retorno entero
        ///</summary>
        ///<param name="fun">Nombre de la función</param>
        ///<param name="datos">Arreglo con los valores de la funcion</param>
        //Ejecucion de funcion con retorno entero
        public int execIntFunction(string fun, string[] datos)
        {
            int retorno = 0;

            try
            {
                this.conectar();
                sql = "select " + fun + " ( ";
                for (int i = 0; i < datos.Length; i++)
                {
                    if (i != 0) { sql += ", "; }
                    sql += "'" + datos[i] + "'";
                }
                sql += " ) as res ; ";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                int r = utl.convertirInt(dtTable.Rows[0]["res"]);
                retorno = r;

                this.desconectar();
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Integer Function - MySQL");
                retorno = -1;
                this.desconectar();
            }

            return retorno;
        }


        ///<summary>
        ///Ejecucion de funcion con retorno objeto
        ///</summary>
        ///<param name="fun">Nombre de la función</param>
        ///<param name="datos">Arreglo con los valores de la funcion</param>
        //Ejecucion de funcion con retorno objeto
        public object execObjectFunction(string fun, string[] datos)
        {
            object retorno = null;

            try
            {
                this.conectar();
                sql = "select " + fun + " ( ";
                for (int i = 0; i < datos.Length; i++)
                {
                    if (i != 0) { sql += ", "; }
                    sql += "'" + datos[i] + "'";
                }
                sql += " ) as res ; ";
                daAdapter = new MySqlDataAdapter(sql, sqlConn);
                dtTable = new DataTable();
                daAdapter.Fill(dtTable);
                retorno = dtTable.Rows[0]["res"];

                this.desconectar();
            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Object Function - MySQL");                
                this.desconectar();
            }

            return retorno;
        }


        ///<summary>
        ///Ejecucion de procedimiento almacenado (retorno booleano con referencia a ejecucion correcta)
        ///</summary>
        ///<param name="procedimiento">Nombre del procedimiento Almacenado</param>
        ///<param name="datos">Arreglo con los valores de la funcion</param>
        //Ejecucion de procedimiento almacenado (retorno booleano con referencia a ejecucion correcta)
        public bool execBoolProcedure(string procedimiento, string[] datos)
        {
            bool retorno = false;

            try
            {
                this.conectar();
                sqlcomm = sqlConn.CreateCommand();
                sqlcomm.CommandType = CommandType.Text;
                sql = "CALL " + procedimiento + " ( ";
                for (int i = 0; i < datos.Count(); i++)
                {
                    if (i != 0) { sql += ", "; }
                    sql += (datos[i] != "") ? "'" + datos[i] + "'" : "''";
                }
                sql += " ) ;";

                sqlcomm.CommandText = sql;
                sqlcomm.ExecuteNonQuery();

                daAdapter = new MySqlDataAdapter("show warnings;", sqlConn);
                DataTable dtw = new DataTable();
                daAdapter.Fill(dtw);

                if (dtw.Rows.Count > 0)
                {
                    retorno = false;
                    foreach (DataRow r in dtw.Rows)
                    {
                        err.AddErrors("Warning: " + utl.convertirString(r["Message"]), "Cod. Error:" + utl.convertirString(r["Code"]), "Procedure - MySQL", sql, this.SystemUser);
                    }
                }
                else
                {
                    retorno = true;
                }

                this.desconectar();

            }
            catch (Exception e)
            {
                err.AddErrors(e, sql, obtenerConnectionString(), "Procedure - MySQL");                
                retorno = false;
                this.desconectar();
            }

            return retorno;
        }
        



        #endregion


    }
}
