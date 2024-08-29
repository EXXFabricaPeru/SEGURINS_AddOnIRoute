using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Integration_IROUTE.Entity
{
    public class Entrega
    {
        [JsonProperty("nroTransporte")] public string nroTransporte { get; set; }
        [JsonProperty("codigoSociedad")] public string codigoSociedad { get; set; }
        [JsonProperty("codigoAlmacen")] public string codigoAlmacen { get; set; }
        [JsonProperty("codigoViajeConsolidado")] public string codigoViajeConsolidado { get; set; }
        [JsonProperty("fechaEstimadaInicioViaje")] public string fechaEstimadaInicioViaje { get; set; }
        [JsonProperty("fechaEstimadaFinViaje")] public string fechaEstimadaFinViaje { get; set; }
        [JsonProperty("nroPedido")] public string nroPedido { get; set; }
        [JsonProperty("nroPedidoWeb")] public string nroPedidoWeb { get; set; } //opcional
        [JsonProperty("fechaEntrega")] public string fechaEntrega { get; set; }
        [JsonProperty("ventanaHorariaInicio")] public string ventanaHorariaInicio { get; set; } //opcional
        [JsonProperty("ventanaHorariaFin")] public string ventanaHorariaFin { get; set; } //opcional
        [JsonProperty("ventanaHorariaInicio2")] public string ventanaHorariaInicio2 { get; set; } //opcional
        [JsonProperty("ventanaHorariaFin2")] public string ventanaHorariaFin2 { get; set; } //opcional
        [JsonProperty("horaCita")] public string horaCita { get; set; } //opcional
        [JsonProperty("tiempoVisita")] public int tiempoVisita { get; set; } //opcional
        [JsonProperty("cantidadBultos")] public double cantidadBultos { get; set; } //opcional
        [JsonProperty("peso")] public double peso { get; set; } //opcional
        [JsonProperty("volumen")] public double volumen { get; set; } //opcional
        [JsonProperty("prioridad")] public string prioridad { get; set; } //opcional
        [JsonProperty("fragil")] public string fragil { get; set; } //opcional
        [JsonProperty("adicional1")] public string adicional1 { get; set; } //opcional
        [JsonProperty("adicional2")] public string adicional2 { get; set; } //opcional
        [JsonProperty("debeCobrar")] public string debeCobrar { get; set; } //opcional
        [JsonProperty("codRastreo")] public string codRastreo { get; set; }  //opcional
        [JsonProperty("viajeCourier")] public int viajeCourier { get; set; } //opcional
        [JsonProperty("cliente")] public cliente Cliente { get; set; }
        [JsonProperty("guias")] public List<guias> Guias { get; set; }
    }

    public class cliente
    {
        [JsonProperty("codigoCliente")] public string codigoCliente { get; set; }
        [JsonProperty("nombre")] public string nombre { get; set; }
        [JsonProperty("telefono")] public string telefono { get; set; } //opcional
        [JsonProperty("correo")] public string correo { get; set; } //opcional
        [JsonProperty("codigoCanal")] public string codigoCanal { get; set; } //opcional
        [JsonProperty("codigoSubcanal")] public string codigoSubcanal { get; set; } //opcional
        [JsonProperty("categoria")] public string categoria { get; set; } //opcional
        [JsonProperty("rubro")] public string codigoSociedad { get; set; } //opcional
        [JsonProperty("sede")] public sede Sede { get; set; }
    }

    public class sede
    {
        [JsonProperty("codigo")] public string codigo { get; set; }
        [JsonProperty("nombre")] public string nombre { get; set; }
        [JsonProperty("direccion")] public string direccion { get; set; }
        [JsonProperty("referencia")] public string referencia { get; set; } //opcional
        [JsonProperty("latitud")] public string latitud { get; set; } //opcional
        [JsonProperty("longitud")] public string longitud { get; set; } //opcional
        [JsonProperty("departamento")] public string departamento { get; set; } //opcional
        [JsonProperty("provincia")] public string provincia { get; set; } //opcional
        [JsonProperty("distrito")] public string distrito { get; set; } //opcional
        [JsonProperty("leadTime")] public string leadTime { get; set; } //opcional
        [JsonProperty("radioControl")] public string radioControl { get; set; } //opcional
        [JsonProperty("tiempoDescarga")] public string tiempoDescarga { get; set; } //opcional
        [JsonProperty("tiempoEspera")] public string tiempoEspera { get; set; } //opcional
        [JsonProperty("contacto")] public contacto Contacto { get; set; } //opcional
    }

    public class contacto
    {
        [JsonProperty("nombre")] public string nombre { get; set; }
        [JsonProperty("apellido")] public string apellido { get; set; } //opcional
        [JsonProperty("telefono")] public string telefono { get; set; } //opcional
        [JsonProperty("docIdentidad")] public string docIdentidad { get; set; } //opcional
        [JsonProperty("correo")] public string correo { get; set; } //opcional
    }

    public class guias
    {
        [JsonProperty("nroGuia")] public string nroGuia { get; set; }
        [JsonProperty("formaPago")] public string formaPago { get; set; } //opcional
        [JsonProperty("pago")] public double pago { get; set; } //opcional
        [JsonProperty("detalleGuia")] public string detalleGuia { get; set; } //opcional
    }
}
