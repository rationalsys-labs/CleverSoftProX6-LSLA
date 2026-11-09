# Demostración · Mayo de 2011

Los nombres comerciales, personas, documentos de identidad, contactos, direcciones y operaciones de este conjunto son ficticios. No proceden de la información operativa de LSLA. Los distritos son referencias geográficas de Lima; `LIMA01` a `LIMA08` son los códigos del catálogo demo.

[Instalación](../LEEME.md) · [Manual de uso](../../MANUAL_DE_USO.md)

## Cuentas de acceso

La clave de todas las cuentas demo es **`Demo2011`**.

| Usuario | Perfil | Acceso |
| --- | --- | --- |
| `ADMIN` | Administrador | Todos los módulos y administración de cuentas. |
| `GEST1` | Administrador | Segunda cuenta para demostrar la administración sin depender de una sola cuenta. |
| `OPER1` | Operador | Los once módulos operativos; no administra usuarios. |
| `REPA1` | Operador | Reparto, consolidados, confirmaciones y Delivery. |
| `INFO1` | Operador | Consulta de documentos, Delivery e informes. |
| `BAJA1` | Operador inactivo | No ingresa; permite demostrar la reactivación desde Usuarios. |

Los checks se editan para operadores. Los administradores tienen acceso completo. Estas cuentas no son cuentas de SQL Server.

## Información incluida

- 7 clientes: cinco habilitados, uno con operaciones inactivas y uno con baja lógica.
- 9 servicios, incluido el servicio general `105`.
- 29 guías, de las que una está anulada; 29 destinatarios.
- 3 grupos y 5 mensajeros, uno de ellos inactivo.
- Historiales de reparto, entregas, ausencias, direcciones incorrectas, rechazos, liberación y reasignación.
- Productos, tipos de documento y servicio, distritos, tarifas demostrativas y correlativos.
- Dos anotaciones pendientes del esquema 2 para exportar actualizaciones.

Las fechas iniciales están en mayo de 2011. Las nuevas operaciones realizadas en la aplicación utilizan la fecha del equipo según el funcionamiento del sistema.

## Casos preparados

| Caso | Referencia | Qué demostrar |
| --- | --- | --- |
| Pedido con capacidad | Servicio `1101`, boleta `1101101` | Consultar cuatro guías y agregar hasta completar seis. |
| Pedido listo para finalizar | Servicio `1102`, boleta `1101102` | Cuatro documentos registrados de cuatro contratados. |
| Pedido finalizado | Servicio `1103`, guía `511010011` | Consultar y reactivar el servicio. |
| Captura manual o Excel | Servicio `1108`, boleta `1101108` | Registrar tres documentos desde cero. |
| Agrupación | Grupo `1`: RUTA COMERCIAL MAYO | Consultar `1101` y `1102`; agrupar otro servicio disponible. |
| Retirada del último servicio | Grupo `3`: PEDIDO ESPECIAL MAYO | Retirar `1107` conservando sus documentos. |
| Nuevo reparto | `511010006`, `511010021`, `511010022` | Formar una lista sin asignaciones previas vigentes. |
| Regeneración de cargo | Cargo `501` | Recuperar tres guías de Andrés Robles - Demo. |
| Liberación | `511010025`, cargo `504` | Liberar y volver a asignar. |
| Guía ya liberada | `511010026` | Volver a repartir sin una asignación vigente. |
| Reasignación e historial | `511010024` | Consultar la entrega posterior a la reasignación. |
| Documento anulado | `511010027` | Comprobar su exclusión del reparto activo. |
| Consolidado nuevo | `511010028` y `511010029` | Registrar la relación y abrir el informe. |
| Confirmación | `511010003` y `511010020` | Registrar entrega o incidencia y consultar historial. |
| Servicio general | `105`, guías `511010001` y `511010002` | Consultar productos diferentes y elegir producto para otra guía. |

## Delivery e informes

Consultar el intervalo **10/05/2011 a 18/05/2011** para el reparto precargado.

| Mensajero | Documentos vigentes | Entregados | Devoluciones | Pendientes |
| --- | ---: | ---: | ---: | ---: |
| Andrés Robles - Demo | 5 | 1 | 2 | 2 |
| César Palacios - Demo | 7 | 4 | 1 | 2 |
| Marina Torres - Demo | 3 | 1 | 1 | 1 |

Son cantidades iniciales: cambian al operar sobre la demo. El intervalo corresponde a la fecha del reparto y se presenta el estado actual de las guías.

En Informe de Status, elegir el origen de estados, seleccionar Código, Empresa, Status y Fecha, y filtrar por una guía conocida. Los otros orígenes permiten explorar clientes, servicios y documentos.

## Ejemplo de importación Excel

Puede descargarse [Importacion.xlsx](Importacion.xlsx), un libro de dos filas listo para el servicio `1108`. La segunda fila deja CONTACTO vacío para demostrar que es opcional.

Para crear otro libro, colocar los cuatro encabezados en la **primera fila**, sin títulos ni filas previas:

| EMPRESA | DIRECCION | DISTRITO | CONTACTO |
| --- | --- | --- | --- |
| Librería Río de Papel - Demo | Av. Muestra Nueva 401 | LIMA02 | Contacto Demo A |
| Oficina Jardín del Norte - Demo | Jr. Muestra Nueva 402 | LIMA05 | |

Guardar como `.xls` o `.xlsx`, seleccionar el servicio `1108`, indicar una descripción e importar. Quedará una plaza para demostrar el registro manual. No copiar al libro la fila de separación de Markdown.

## Actualizaciones

Instalar dos bases demo desde los mismos scripts, por ejemplo `DCSP_Demo` y `DCSP_DemoDestino`, para comenzar con identidades compatibles. Exportar desde la primera y cargar el archivo en la segunda. Repetir la importación permite comprobar que no se duplican las filas.

La carga demo exige tablas vacías. Para comenzar otro recorrido sin conservar las operaciones realizadas, instalarla con un nombre de base nuevo.
