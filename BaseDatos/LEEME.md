# Base de datos de Clever SoftPro X6 · LSLA

Scripts completos para crear una base nueva: tablas, claves primarias y foráneas, índices, restricciones, vistas, procedimientos y protección de direcciones utilizadas por documentos. La información de demostración se carga por separado.

[Presentación](../README.md) · [Manual de uso](../MANUAL_DE_USO.md) · [Datos demo](Demo/LEEME.md)

## Preparación

- SQL Server y una cuenta de Windows autorizada para crear la base e instalar sus objetos.
- `sqlcmd` para los instaladores. Los archivos SQL individuales también pueden consultarse en SQL Server Management Studio.
- Ejecutar los comandos desde la **raíz del repositorio**: las directivas `:r` utilizan esa ubicación.
- Elegir un nombre de base nuevo. Los instaladores no eliminan ni reemplazan bases existentes.

La aplicación usa autenticación de Windows por defecto. La cuenta de Clever es independiente del acceso de Windows a SQL Server; no se necesita habilitar `sa`.

## Instalación completa con demostración

```powershell
sqlcmd -S 'localhost\SQLEXPRESS' -E -b -f 65001 -v NombreBase="DCSP_Demo" -i BaseDatos\InstalarDemo.sql
```

Este comando crea la base, instala todos los objetos, verifica la estructura y carga los datos ficticios de mayo de 2011. Permite iniciar sesión con `ADMIN` / `Demo2011`.

El nombre se elige mediante `NombreBase`: letras, números y guion bajo. Para una instalación nueva con el nombre predeterminado de la aplicación, puede utilizarse `NombreBase="DCSP"`, siempre que esa base no exista.

Con `sqlcmd` y un certificado local de SQL Server puede ser necesario añadir `-C`. Esa opción pertenece al cliente de instalación; no cambia la configuración de la aplicación.

## Instalar solamente la estructura

```powershell
sqlcmd -S 'localhost\SQLEXPRESS' -E -b -f 65001 -v NombreBase="DCSP_Estructura" -i BaseDatos\Instalar.sql
```

La base queda vacía, sin cuentas ni catálogos. Para cargar luego la demostración:

```powershell
sqlcmd -S 'localhost\SQLEXPRESS' -E -b -f 65001 -v NombreBase="DCSP_Estructura" -i BaseDatos\CargarDemo.sql
```

La carga comprueba que **todas las tablas estén vacías**. No borra datos, no reinicia contadores y no mezcla información con una instalación en uso. Ante un error de datos, revierte su transacción completa.

## Orden de los scripts

Los instaladores ejecutan esta secuencia automáticamente:

| Archivo | Contenido |
| --- | --- |
| `00_CrearBase.sql` | Base, intercalación y nivel de compatibilidad. |
| `01_Estructura.sql` | Tablas, claves, relaciones, índices, valores predeterminados y validaciones. |
| `02_Procedimientos.sql` | Operaciones generales, correlativos y protección de direcciones históricas. |
| `03_Vistas.sql` | Consultas de formularios, búsquedas, historial e informes. |
| `04_Servicios.sql` | Cantidad, finalización, reactivación y anulación. |
| `05_Grupos.sql` | Grupos y pertenencia de servicios. |
| `06_Reparto.sql` | Asignación, cargos y liberación. |
| `07_Estados.sql` | Confirmaciones por lote y Delivery. |
| `08_InformesConfiguracion.sql` | Estados e intercambio de actualizaciones. |
| `09_Usuarios.sql` | Usuarios, perfiles, permisos y autenticación. |
| `10_Verificar.sql` | Comprobación estructural de claves, relaciones, vistas y restricciones. |
| `Demo/01_Datos.sql` | Catálogos y datos ficticios mediante INSERT. |
| `Demo/02_Verificar.sql` | Cantidades, fechas y coherencia de la demostración. |

La estructura final contiene **21 tablas**. `Instalar.sql` termina después de la comprobación estructural; `InstalarDemo.sql` continúa con los datos. `CargarDemo.sql` utiliza una estructura vacía ya instalada.

La comprobación estructural no crea operaciones de negocio. Puede repetirse con `sqlcmd -d NOMBRE_BASE -i BaseDatos\10_Verificar.sql`.

## Conectar la aplicación

Copiar `Config/Clever.example.ini` a `Config/Clever.local.ini`, junto al ejecutable:

```ini
[Conexion]
Servidor=localhost\SQLEXPRESS
Base=DCSP_Demo
AutenticacionWindows=1
Usuario=
Clave=
```

La carpeta `Config` se resuelve desde la ubicación del ejecutable. El archivo local no se incluye en Git. Sin él, la conexión predeterminada es `localhost\SQLEXPRESS`, base `DCSP`.

Se requieren VB6, ADO, `MSCOMCTL.OCX`, `MSFLXGRD.OCX`, `MSCOMCT2.OCX`, `COMDLG32.OCX` y `HookMenu.ocx`. El consolidado utiliza `MSDBRPTR.DLL` y `msstdfmt.dll`. Los componentes COM deben estar registrados en Windows. Los instaladores y licencias se obtienen de sus medios correspondientes.

Excel debe estar instalado para importar o generar libros. La aplicación se compila como programa de 32 bits.

## Reglas de los datos

- Clientes, servicios, documentos, usuarios y mensajeros se relacionan mediante claves numéricas.
- La dirección del destinatario utilizada por un documento conserva su identidad histórica.
- El último estado se determina por fecha y después por número de evento.
- Los cargos numerados usan el correlativo `CRA`; los consolidados utilizan `Ranche=0`.
- Las bajas conservan las relaciones y el historial.
- El perfil `0` es Administrador y el `1` es Operador. Los permisos del operador se asignan por cuenta y módulo.
- Las claves se almacenan como MD5. Las credenciales demo son públicas y pertenecen exclusivamente al conjunto ficticio.
- Los archivos del esquema 2 no incluyen usuarios, claves, mensajeros ni todos los catálogos. El intercambio requiere identidades compatibles y no resuelve conflictos entre bases independientes.

## Compatibilidad

Se utiliza nivel de compatibilidad **100** y construcciones de SQL Server 2008. La instalación se comprueba con SQL Server 2025 en ese nivel; no se certifica un motor SQL Server 2008 real.

Las funciones del sistema, Excel y Data Report se verifican por separado de la salida física de la impresora.
