# Manual de uso · Clever SoftPro X6 — LSLA

Sistema para el registro, distribución y seguimiento de documentos de un courier local.

[Presentación](README.md) · [Recorrido con datos demo](BaseDatos/Demo/LEEME.md) · [Instalación](BaseDatos/LEEME.md)

## 1. Cómo se organiza el trabajo

El **cliente** encarga un **servicio**: un pedido con boleta, producto y cantidad. Cada documento del pedido tiene una **guía** que permite identificarlo durante toda la operación.

Los documentos se asignan a un **mensajero** para su reparto. El cargo o consolidado reúne esa lista. Después se registra el resultado: entregado, ausente, dirección incorrecta u otra incidencia. Las consultas e informes permiten conocer el avance.

Un **grupo** reúne servicios para consultarlos juntos. Finalizar un servicio significa terminar su registro de documentos; no equivale a indicar que todas las entregas ya se realizaron.

## 2. Ingreso y navegación

Seleccionar el usuario, escribir su clave y acceder. Los menús disponibles dependen de sus permisos. La barra superior ofrece accesos a las funciones principales y las barras inferiores muestran información de la operación y la selección.

En la demostración puede utilizarse `ADMIN` con la clave `Demo2011`. Sus casos preparados están en el [recorrido demo](BaseDatos/Demo/LEEME.md).

Utilizar **Cancelar**, **Nueva Lista** o **Volver**, según la pantalla, para abandonar la operación actual. Leer la confirmación antes de descartar datos pendientes. Para cerrar el sistema, utilizar **Archivo → Salir**.

## 3. Mantenimiento de Clientes

Abrir **Operaciones → Mantenimiento de Clientes** o el botón **Clientes**.

- Para consultar, filtrar la lista y seleccionar un cliente.
- Para crear, elegir **Crear Cliente**, completar la ficha y guardar.
- Para modificar, seleccionar primero el cliente y elegir la modificación.
- **Cancelar** permite volver a la consulta sin guardar el borrador.
- La baja conserva las relaciones e historial. Revisar código, nombre y documento en la confirmación.

La ficha requiere nombre, documento de identidad, dirección, distrito, representante, contacto y teléfono general. Conviene completarla antes de registrar pedidos.

## 4. Visualizar Información de Registros

Abrir la consulta desde **Operaciones** o el botón **Registros**.

Escribir el **código de la guía o documento** y pulsar **Enter**. No es el código del servicio ni una búsqueda por nombre. La lupa abre un selector para localizar y escoger un documento.

La pantalla presenta servicio, cliente, destinatario, dirección, producto, descripción e historial. Si el servicio está finalizado, puede reactivarse para continuar el registro cuando corresponda.

## 5. Generar Registros y Etiquetas

Abrir **Operaciones → Crear Registros de Reparto** o el botón **Etiquetas**.

### Servicio

Elegir la creación de servicio y seleccionar el cliente con **F9**. Indicar boleta, tipo de servicio, producto y cantidad contratada. Guardar el pedido y revisar su código.

La boleta es numérica, positiva y única; la cantidad es un entero positivo. Seleccionar un servicio existente antes de registrar sus documentos. La pantalla muestra cantidad registrada y capacidad disponible.

### Registro manual

Elegir la creación de registro. Indicar la referencia del cliente o marcar **Sin Data del Cliente** si no dispone de ella. Completar destinatario, dirección y distrito. Contacto y descripción manual son opcionales. Guardar y continuar con el siguiente documento.

La guía generada será su identificador para consultas, reparto y confirmaciones. Conservar el código mostrado cuando se informe que el documento quedó guardado.

### Importación Excel

Preparar una hoja con **EMPRESA, DIRECCION, DISTRITO y CONTACTO** en la primera fila. DISTRITO contiene el código del catálogo utilizado por esa instalación. CONTACTO puede quedar vacío.

Seleccionar el servicio y el archivo `.xls` o `.xlsx`, indicar una descripción y cargar. El lote debe caber en la capacidad disponible. Si se informa un error, corregir el archivo antes de repetir la carga. No agregar títulos decorativos ni dejar filas incompletas entre los registros.

El [recorrido demo](BaseDatos/Demo/LEEME.md#ejemplo-de-importación-excel) incluye un ejemplo de dos filas.

### Cantidad, finalización y anulación

La cantidad puede ajustarse con el servicio abierto y nunca por debajo de los documentos registrados. Finalizar se utiliza cuando la cantidad registrada coincide con la contratada. Para continuar después, reactivar desde la consulta de un documento.

Anular requiere seleccionar una guía de un servicio abierto. Conserva el registro para su historial y evita que continúe como documento activo.

El servicio general **105** permite registrar sin el límite habitual de cantidad y requiere elegir el producto correspondiente. No se finaliza como un servicio ordinario.

### Impresión

Elegir la modalidad indicada en la pantalla. La reimpresión recupera documentos ya registrados; no crea otros. En regeneración masiva se utiliza el código de guía; **Volver** recupera el trabajo anterior.

Si el mensaje indica que los registros quedaron guardados pero no se pudo imprimir, consultar o reimprimir esas guías. No repetir su creación.

## 6. Gestión de Grupos de Servicios

Abrir **Operaciones → Gestionar Grupos de Servicios**.

Crear un grupo con nombre identificable o buscar uno existente. El selector permite precisar la búsqueda con **F3**. Marcar servicios disponibles y agruparlos. Utilizar **Cargar Datos del Grupo** para consultar sus documentos y resúmenes.

Los filtros reducen los documentos visibles; Excel exporta la consulta mostrada. Para retirar un servicio, seleccionarlo y pulsar **Suprimir**: solo se retira su pertenencia al grupo, conservando los documentos. También puede retirarse el último servicio.

Cancelar permite descartar una selección pendiente de agrupación.

## 7. Registro de Reparto y Ranches

Abrir el control de reparto desde **Operaciones → Control de Datos de Reparto** o el botón **Cargos**.

Ingresar las guías y pulsar **Enter** para agregarlas. Elegir el mensajero antes de registrar. Para retirar una fila pendiente, seleccionarla y pulsar **Suprimir**.

Revisar y registrar el reparto. El sistema asigna un número de cargo y permite imprimirlo. **Nueva Lista** limpia el trabajo de la pantalla; no deshace un reparto guardado.

Para reimprimir, activar la regeneración, ingresar el cargo y pulsar Enter. **Último Cargo** recupera el último disponible en esa ventana y **Volver** abandona la consulta.

Una guía asignada debe liberarse antes de repartirse nuevamente. Revisar destinatario, mensajero, cargo y fecha en la confirmación. Después, volver a ingresar la guía con Enter para agregarla a la nueva lista. La liberación conserva el documento y sus eventos anteriores.

## 8. Generar Consolidados

Elegir mensajero y agregar guías con **Enter**. Revisar la relación y registrar para abrir el informe de consolidado. Esta modalidad no genera un número de cargo.

**Último Informe** reabre el último consolidado mientras esa ventana permanezca abierta. **Nueva Lista** inicia otra relación. Si el reparto quedó guardado y falla la apertura del informe, reintentar su generación sin registrar otra vez las mismas guías.

## 9. Confirmación de Registros

Abrir **Registrar Status** o el botón **Status**.

Elegir estado, indicar fecha y marcar la casilla de hora. Escribir la observación si corresponde. Ingresar la guía y pulsar **Enter** para agregarla al lote pendiente. La medianoche es válida cuando se selecciona expresamente.

Cada fila conserva los datos con los que se agregó. Revisar historial y lista antes de registrar el lote. **Suprimir** retira una confirmación pendiente; **Cancelar** permite descartar los datos no registrados.

La vista previa distingue lo pendiente de lo guardado. Si falla solo la consulta del historial, utilizar **Ver Historial**. Cuando el mensaje confirme el guardado, no registrar de nuevo por ese fallo de consulta.

## 10. Control de Documentos a Delivery

Seleccionar mensajero y fechas **Desde** y **Hasta**, y pulsar **Cargar**.

Se reúnen las asignaciones vigentes dentro del intervalo de reparto, mostrando el estado actual y cantidades entregadas, devueltas y pendientes. No es una fotografía del estado que tenían las guías al terminar una fecha antigua.

Cambiar filtros exige cargar de nuevo. **Limpiar Consulta** reinicia la selección; **Generar Excel** exporta los resultados.

Para los datos demo, seleccionar las fechas de mayo de 2011 de su recorrido, en lugar de dejar las fechas actuales ofrecidas inicialmente.

## 11. Informe de Status

Abrir **Formularios e Informes → Informe de Status**.

Elegir origen: estados, servicios, clientes o documentos. Marcar las columnas deseadas y completar, si hace falta, hasta cuatro filtros con campo, condición y valor.

Las conexiones **Y/O** se evalúan de izquierda a derecha. Utilizar Y cuando deban cumplirse ambas condiciones y O cuando se admita cualquiera. Indicar fechas válidas para campos de fecha y valores numéricos para cantidades y códigos numéricos.

Generar Excel, guardar texto o imprimir. Revisar los filtros antes de producir una salida extensa.

## 12. Configuración de Informaciones

Abrir **Configuración → Configurar Variables de Sistema**.

Crear un estado o seleccionar uno existente para modificarlo. La baja conserva los eventos que utilizaron ese estado. **Reparto** y **Entregado** son estados básicos que deben permanecer activos. Cancelar descarta un borrador.

La opción de distritos permite consultar su catálogo y códigos; no ofrece edición de distritos.

## 13. Actualizaciones

Abrir **Configuración → Actualizaciones**.

Para generar un archivo, elegir destino y guardar los pendientes. Para cargarlo, seleccionar la modalidad de importación, buscar el archivo y confirmar. Ante un error se revierten los cambios del lote; reaplicar el mismo archivo no agrega copias de sus filas.

El intercambio requiere instalaciones compatibles. No transporta cuentas ni todos los catálogos y no combina automáticamente bases que hayan trabajado de forma independiente.

## 14. Administración de Usuarios

Con una cuenta administradora, abrir **Usuarios → Administración de Usuarios**.

- **Información Interna de Usuarios** vuelve a la consulta desde Crear o Modificar.
- **Actualizar Lista** vuelve a consultar los registros.
- **Crear Usuario** recibe nombres, apellidos, usuario, perfil, correo opcional y clave confirmada.
- **Modificar Usuarios** permite cambiar datos, perfil y situación de una cuenta seleccionada, conservando su identificador.
- **Cambiar Clave** establece una clave nueva confirmada.
- **Dar de Baja** impide un nuevo ingreso y conserva las referencias históricas.
- **Cancelar** abandona los datos pendientes y vuelve a consulta.

El usuario admite hasta cinco letras o números y las claves nuevas requieren de 6 a 150 caracteres.

El **Administrador** tiene acceso completo y gestiona cuentas; sus checks están marcados y bloqueados. El **Operador** recibe permisos por módulo y por cuenta: dos operadores pueden tener accesos diferentes. Marcar los módulos necesarios y pulsar **Guardar Permisos**; Cancelar no aplica la selección.

La cuenta administradora de la sesión está protegida para evitar su baja o pérdida de administración durante el trabajo.

## 15. Recorrido sugerido

1. Ingresar como ADMIN y consultar clientes.
2. Abrir una guía conocida y revisar su servicio e historial.
3. Registrar documentos en un pedido con capacidad.
4. Consultar un grupo y su resumen.
5. Preparar un reparto y recuperar un cargo existente.
6. Generar un consolidado.
7. Confirmar una entrega y revisar el historial.
8. Consultar Delivery y generar un informe.
9. Explorar configuración, cuentas y permisos.
10. Exportar actualizaciones y probarlas en otra base demo compatible.

El [recorrido demo](BaseDatos/Demo/LEEME.md) proporciona las cuentas, fechas y códigos para cada paso.
