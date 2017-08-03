using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SincronizadorCompensatoriosCCO.Modelo
{
    class MotivosCompensatorios
    {
        private List lstMotivosCompensatorios;
        private ClientContext context;


        public MotivosCompensatorios(ClientContext context, string NombreLista)
        {

            this.context = context;
            this.lstMotivosCompensatorios = this.context.Web.Lists.GetByTitle(NombreLista);
        }

        public bool GuardarMotivosCompensatorios(
                            int IdMotivo,
                            string MotivoDelCompensatorio,
                            string NombreDeUsuario,
                            decimal NumeroDeDias

        )
        {
            try
            {

                Principal UsuarioPersona = this.context.Web.SiteUsers.GetByEmail(NombreDeUsuario);
                context.Load(UsuarioPersona);
                this.context.ExecuteQuery();

                CamlQuery query = new CamlQuery();
                query.ViewXml = "<View>" +
                  "<Query>" +
                    "<Where>" +
                        "<And>" +
                              "<Eq>" +
                                "<FieldRef Name='IdMotivo'/>" +
                                "<Value Type='Lookup'>" + IdMotivo + "</Value>" +
                              "</Eq>" +
                              "<Eq>" +
                                "<FieldRef Name='NombreUsuario' LookupId='TRUE'/>" +
                                "<Value Type='Lookup'>" + UsuarioPersona.Id + "</Value>" +
                              "</Eq>" +
                        "</And>" +
                    "</Where>" +
                  "</Query>" +
                  "</View>";

                ListItemCollection itemsCompensatorios = lstMotivosCompensatorios.GetItems(query);
                context.Load(itemsCompensatorios);
                context.ExecuteQuery();
                if (itemsCompensatorios.Count > 0)
                {
                    foreach (ListItem Compensatorios in itemsCompensatorios)
                    {
                        Compensatorios["NumeroDias"] = Convert.ToDecimal(Compensatorios["NumeroDias"]) + NumeroDeDias;
                        Compensatorios.Update();
                        this.context.ExecuteQuery();
                    }
                }
                else
                {
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem newItem = this.lstMotivosCompensatorios.AddItem(itemCreateInfo);
                    newItem["IdMotivo"] = IdMotivo;
                    newItem["MotivoCompensatorio"] = MotivoDelCompensatorio;
                    newItem["NombreUsuario"] = UsuarioPersona.Id;
                    newItem["NumeroDias"] = NumeroDeDias;

                    newItem.Update();
                    this.context.ExecuteQuery();
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}
