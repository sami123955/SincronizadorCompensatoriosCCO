using System;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SincronizadorCompensatoriosCCO.Modelo
{
    public class RegistrosExcel
    {

        private List lstRegistrosExcel;
        private ClientContext context;


        public RegistrosExcel(ClientContext context, string RegistroExcel)
        {

            this.context = context;
            this.lstRegistrosExcel = this.context.Web.Lists.GetByTitle(RegistroExcel);
        }

        public bool GuardarRegistroExcel(
                            int Id_Motivo,
                            string MotivoDelCompensatorio,
                            string NombreDeUsuario,
                            decimal NumeroDeDias,
                            string ResponsableCarga,
                            string Observacion,
                            int IdCargaExcel,
                            string fechaCompensantorio

        )
        {
            try
            {



                Principal UsuarioPersona = this.context.Web.SiteUsers.GetByEmail(NombreDeUsuario);
                context.Load(UsuarioPersona);
                this.context.ExecuteQuery();

                Principal usuarioResponsable = this.context.Web.SiteUsers.GetByEmail(ResponsableCarga);
                context.Load(usuarioResponsable);
                this.context.ExecuteQuery();

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = this.lstRegistrosExcel.AddItem(itemCreateInfo);
                newItem["IdMotivo"] = Id_Motivo;
                newItem["IdArchivo"] = IdCargaExcel;
                newItem["MotivoCompensatorio"] = MotivoDelCompensatorio;
                newItem["NombreUsuario"] = UsuarioPersona.Id;
                newItem["NumeroDias"] = NumeroDeDias;
                newItem["ResponsableCarga"] = usuarioResponsable.Id;
                newItem["Observacion"] = Observacion;
                newItem["FechaCompensatorio"] = fechaCompensantorio;
                newItem.Update();
                this.context.ExecuteQuery();

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}

