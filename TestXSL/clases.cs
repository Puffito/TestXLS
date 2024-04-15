﻿using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Security.Cryptography;

namespace TestXLS
{
	public class SyIntegradores
	{
        public List<SyIntegrador> Listado;

		public SyIntegradores()
		{
            Listado = new List<SyIntegrador>();
        }
	}

    public class SyIntegrador
    {
        private string xnombre;
        private string xtipo;
        private string xequipos;
        private string xindice;
        private string xindice2;

        public SyIntegrador(string nombre, string tipo, string equipos, string indice, string indice2)
        {
            xnombre = nombre;
            xtipo = tipo;
            xequipos = equipos;
            xindice = indice;
            xindice2 = indice2;
        }

        public string nombre
        {
            get
            {
                return xnombre;
            }
            private set
            {
                xnombre = value;
            }
        }
        public string tipo
        {
            get
            {
                return xtipo;
            }
            private set
            {
                xtipo = value;
            }   
        }
        public string equipos
        {
            get
            {
                return xequipos;
            }
            private set
            {
                xequipos = value;
            }
        }
        public string indice
        {
            get
            {
                return xindice;
            }
            private set
            {
                xindice = value;
            }
        }
        public string indice2
        {
            get
            {
                return xindice2;
            }
            private set
            {
                xindice2 = value;
            }
        }
    }
}
