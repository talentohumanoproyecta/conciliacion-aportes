# Conciliación de aportes – Versión 1

Aplicativo web que cruza, **trabajador por trabajador**, la nómina contabilizada (auxiliar 2370/2380 de Siigo o Mekano) contra la planilla PILA del operador (Aportes en Línea o Mi Planilla) **antes de pagar**, y explica cada diferencia con su responsable.

> 🔒 Los archivos se procesan en el navegador de quien usa la página. **No se suben a ningún servidor**, ni a GitHub. El repositorio solo contiene el código.

---

## 1. Qué detecta

| Semáforo | Caso | Ejemplo real |
|---|---|---|
| 🔴 | Entidad distinta (EPS, AFP, CCF) entre contabilidad y planilla | Pastrana: Sanitas → Nueva EPS Movilidad |
| 🔴 | Clase de riesgo ARL distinta entre contabilidad y planilla | Gladys Canasto: riesgo 5 vs riesgo 1 |
| 🔴 | Riesgo de la planilla distinto al del catálogo para ese centro | — |
| 🔴 | Valor distinto (IBC, días, novedades) por encima de la tolerancia | Ingrid Medina: caja de compensación +$27.300 |
| 🔴 | Trabajador en contabilidad que no está en ninguna planilla cargada, o al revés | Rubén Girón (planilla Profesionales) |
| 🔴 | Entidad mal configurada en el software (a nombre del propio trabajador) | Luis Carlos Galvis |
| 🟡 | Misma entidad con otro código (EPS037 / EPS041) | 9 trabajadores de San Vicente |
| 🟡 | Centro de trabajo que no está en el catálogo | — |
| ✅ | Redondeo dentro de la tolerancia | Por defecto $1.000 por trabajador y $5.000 por entidad |

## 2. Archivos que acepta (tal cual salen, sin organizarlos)

| Archivo | De dónde sale | Quién lo sube |
|---|---|---|
| Auxiliar contable 2370/2380 (.xlsx) | Siigo | Contabilidad |
| Auxiliar contable (.xls) | Mekano | Contabilidad |
| Plano PILA (.txt), opcional | Siigo; sirve para identificar cédulas y códigos de EPS | Contabilidad o Talento Humano |
| Detalle de planilla (.xlsx), **sin pagar**; se pueden cargar varias | Aportes en Línea / Mi Planilla | Talento Humano |
| Catálogo de centros de trabajo (.xlsx) | Carpeta `Catalogo` en Teams | Se carga una vez |

## 3. Uso mensual

1. **Talento Humano** carga la planilla en el operador **sin pagar**, descarga el detalle y lo guarda en `Aportes/AAAA-MM/Empresa/`.
2. **Contabilidad** exporta el auxiliar y lo guarda en la misma carpeta.
3. Se abre la pestaña **Conciliación** en el canal de Teams, se arrastran los archivos y se pulsa **Conciliar**.
4. **Talento Humano** corrige lo marcado en rojo y amarillo (en el operador o en la ficha del trabajador). Si cambió la planilla, se repite el paso 3.
5. Se descarga el **informe Excel** y se guarda en la carpeta del mes. Luego se paga la planilla.

## 4. Catálogo: cambiar un riesgo o crear un centro

Hay una sola fuente oficial: `Catalogo/Catalogo_Centros_Trabajo.xlsx` en Teams.

- **Desde el aplicativo:**
  1. Abre **Ver / editar centros**.
  2. Cambia el riesgo o pulsa **+ Agregar centro**.
  3. Pulsa **Descargar catálogo** y reemplaza el archivo en la carpeta `Catalogo` de Teams.
- **Cuando aparece un centro nuevo en la planilla:** la pestaña **Centros nuevos** permite asignarle tipo y riesgo con un clic. Después se descarga el catálogo igual que arriba.
- **Directamente en Excel:** abre el archivo desde Teams, edita la fila y guarda. Quienes ya lo tengan cargado deben pulsar **Cargar catálogo** de nuevo.
- **Control:**
  - Solo 2 personas con permiso de edición en la carpeta `Catalogo`.
  - SharePoint guarda el historial de versiones (quién cambió qué y cuándo).
  - Para restaurar una versión: clic derecho sobre el archivo → **Historial de versiones**.

## 5. Publicar en GitHub Pages (una sola vez, unos 10 minutos)

1. Crea una cuenta en <https://github.com> con el correo corporativo.
2. Pulsa **New repository**. Nombre: `conciliacion-aportes`. Tipo: **Public** (GitHub Pages gratis lo exige; solo contiene código, **nunca datos**). Pulsa **Create repository**.
3. En el repositorio, pulsa **Add file → Upload files**. Arrastra **solo** `index.html`, la carpeta `js` (con `engine.js` y `app.js`), `README.md` y `.gitignore`. Pulsa **Commit changes**.
4. Ve a **Settings → Pages**. En *Source* elige **Deploy from a branch**, rama `main`, carpeta `/ (root)`, y pulsa **Save**.
5. En 1 o 2 minutos la página queda en `https://TU-USUARIO.github.io/conciliacion-aportes/`.

> ⚠️ Nunca subas al repositorio auxiliares, planillas ni el catálogo. El archivo `.gitignore` lo evita si usas Git, pero al subir por la web hay que tener cuidado.

Para actualizar una versión: **Add file → Upload files** con los archivos nuevos (reemplazan a los anteriores).

## 6. Montar en Teams y SharePoint (sin necesidad de administrador)

1. **Crear el equipo:** en Teams pulsa **Equipos → Crear equipo → Desde cero → Privado**. Nombre: *Conciliación Aportes*. Agrega a Contabilidad y Talento Humano. Así se crea también el sitio de SharePoint.
2. **Crear las carpetas:** en el canal *General*, pestaña **Archivos**, crea esta estructura:
   ```
   Aportes/
     2026-11/
       San Vicente/
       Proyecta Administración/
       ...
   Catalogo/
     Catalogo_Centros_Trabajo.xlsx
   ```
3. **Agregar la pestaña de la aplicación:** pulsa **+** (Agregar una pestaña) → **Sitio web**. Nombre: *Conciliación*. URL: la de GitHub Pages.
4. **Permisos del catálogo:** en SharePoint, abre la carpeta `Catalogo` → **Administrar acceso**. Deja la edición solo a 2 personas y lectura al resto.

**El usuario administrador solo será necesario en la fase 2**, cuando la aplicación lea y guarde directamente en SharePoint sin descargar archivos.

## 7. Estructura del código

```
index.html      interfaz
js/engine.js    lectores de archivos y motor de cruce (sin interfaz)
js/app.js       comportamiento de la página
```

- Librería externa: SheetJS 0.18.5 desde cdnjs, para leer `.xls`, `.xlsx` y escribir el informe.
- Cuentas que se leen: 237005 salud, 237006 ARL, 237010 caja de compensación (SENA e ICBF se ignoran), 238030 pensión.
