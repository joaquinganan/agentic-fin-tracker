# Agentic Finance Tracker — contexto para trabajar en el proyecto con Claude

> **Este archivo es el contexto compartido del proyecto.** Claude Code lo carga automáticamente al abrir el
> repositorio; en una conversación de Claude, adjúntalo (o agrégalo al conocimiento de un Proyecto) y pide que lo lea
> antes de cualquier cambio. Está al día con la **v1.1.63** (5 de octubre de 2026). **Se actualiza en cada versión,
> junto con el CHANGELOG** (ver sección 10). El código, el historial de cada cambio y las guías de uso están en el
> repositorio; este archivo cubre lo que **no** está ahí: las reglas de trabajo, las trampas conocidas y los pendientes.
> Las preferencias personales de cada uno (rutas locales, forma de trabajar) van en `CLAUDE.local.md`, que no se sube.

**Repositorio:** https://github.com/joaquinganan/agentic-fin-tracker

| Documento del repo | Para qué |
|---|---|
| `docs/GUIA.md` | Guía de uso en español (instalación, Setup Wizard, día a día, inversiones, problemas comunes). |
| `docs/USER_GUIDE.md` | La misma guía en inglés. |
| `CHANGELOG.md` | **Fuente de verdad del historial**: qué cambió en cada versión y por qué. Léelo antes de tocar algo relacionado. |
| `tests/README.md` | Qué cubre cada archivo de pruebas. |

---

## 1. Qué es

Un sistema en **Google Apps Script** pegado a una hoja de Google Sheets. Lee las alertas bancarias de Gmail (bancos
dominicanos), registra cada consumo y transferencia en la hoja **Transactions**, categoriza, arma un **Dashboard**, y
envía un resumen diario y otro mensual por correo. También lleva inversiones (órdenes y dividendos de HAPI desde sus
correos, fondos y pensión por saldo), lee estados de cuenta en PDF (Banesco) y tiene un panel **📘 Start here** que
verifica la configuración. Es de un solo usuario por hoja: cada persona instala su propia copia en su cuenta de Google.

## 2. Estado actual (v1.1.63)

- **235 pruebas** en Node (sin dependencias), que corren dos veces: sobre `src/` y sobre el archivo único `dist/`.
- **CI** (GitHub Actions, Node 20 y 22): pruebas, verificación de que `dist/` corresponde a `src/`, y pruebas sobre `dist/`.

| Archivo (`src/`) | Responsabilidad |
|---|---|
| `01_main.gs` | Menú, corridas (diaria y por rango de fechas), Setup Wizard, configuración, disparadores. |
| `02_categorizer.gs` | Categorías por comercio, palabras clave, paleta de colores. |
| `03_gmailMonitor.gs` | Búsqueda en Gmail, `BANK_PATTERNS` (un bloque por banco), extractores por formato, filtros (promociones, no transaccionales, rechazados), lectura de estados de cuenta. |
| `04_sheetsWriter.gs` | Guardado en Transactions, recategorización, hojas derivadas (Bank Transfers, Incoming Transfers, Raw_<BANCO>), Dashboard, estilos, Unrecognized, catálogo de tarjetas (`CARD_PRODUCTS`). |
| `05_dailySummary.gs` / `06_monthlySummary.gs` | Correos de resumen diario y mensual. |
| `07_investments.gs` | Ledger, Holdings, Portfolio History, rendimiento, HAPI, formularios de saldos y de posiciones. |
| `08_startHere.gs` | Panel 📘 Start here (lista de pasos que se verifica sola = health check). |
| `09_pdfText.gs` | Lector de PDF en JavaScript puro (inflate, fuentes, posiciones de texto). Sin Drive API. |

## 3. Instalar y actualizar

- Se instala pegando **un solo archivo**, `dist/FinancialTracker.gs`, como único archivo de código del proyecto de Apps
  Script (Extensions › Apps Script). La línea 2 dice la versión.
- Actualizar = seleccionar todo en ese archivo, pegar la versión nueva y guardar.
- `dist/` se genera con `npm run bundle` a partir de `src/`. **Nunca editar `dist/` a mano.**

## 4. Bancos y formatos soportados

| Banco | Qué lee | Remitente |
|---|---|---|
| LAFISE | Consumos (dos plantillas), pagos de tarjeta, transferencias enviadas (app y "Aviso de transferencia en banco local" de la banca en línea: los fallidos se descartan; a una cuenta propia, Exclude), transferencias entrantes (Pagos al Instante) | bancolafise.com, digital@ y bancanet@notificaciones.lafise.com, notificacioneslafisedo@lafise.com.do, PagosAlInstanteMT103@lafise.com |
| BANESCO | Consumos, transferencias enviadas y recibidas (v1.1.62, sin muestra propia), **estado de cuenta de ahorros en PDF** (todos los créditos) | notificaciones@banesco.com.do, estadodecuenta@banesco.com.do |
| BHD | Consumos (tabla con varias filas), transferencias | bhd.com.do |
| POPULAR | Consumos, Código Cash | popularenlinea.com |
| BDI | Consumos (tabla, solo filas APROBADA), transferencias interbancarias enviadas ("[Salida]") y recibidas ("Recibida"), pago de tarjeta desde la cuenta | bdi.com.do |
| SCOTIABANK | Autorizaciones con tarjeta, Pagos al Instante recibidos | alertas@scotiabank.com |
| QIK | Consumos con tarjeta de crédito, retiros con Código CASH | notificaciones@qik.do, no-reply-qik@qik.com.do |
| BANRESERVAS | Transferencias enviadas (con impuesto aparte), retiros TuEfectivo, transferencias recibidas | NotificacionesTuBancoApp@banreservas.com, notificaciones@banreservas.com |
| HAPI (inversiones) | Órdenes (mercado y límite), dividendos, avisos de depósito | hapi.trade |

`BANK_ORDER` (en `03_gmailMonitor.gs`) es **la única lista de bancos**: las casillas del Wizard, View Config y el
orden de pestañas salen de ella.

**Bancos estrictos** (`strict: true`: BDI, SCOTIABANK, QIK, BANRESERVAS): un correo que ningún extractor del banco
reconoce va a Unrecognized, nunca al adivinador genérico (`extractAllAmounts`). Un banco puede leer el tipo desde el
cuerpo con `detectType(subject, flatText)` cuando el asunto no lo dice (Banreservas) o es igual para enviado y recibido (BDI).

---

## 5. Cómo se entrega cada cambio

1. Cambiar `src/`, agregar o ajustar pruebas, `npm test`.
2. Subir la versión en **tres lugares**: `SCRIPT_VERSION` en `src/01_main.gs`, `package.json` y una entrada en
   `CHANGELOG.md` (en inglés, explicando el porqué). Agregar una fila en `tests/README.md` y un archivo
   `tests/v1NN.test.js` con las pruebas de esa versión.
3. `npm run bundle`, luego `npm run bundle:check` y `npm run test:bundle`.
4. Generar el **patch contra el estado actual de GitHub**: clonar el repo, copiar la versión nueva encima,
   `git add -A`, `git diff --cached --binary > archivo.patch`. Verificarlo aplicándolo en **otro clon limpio**:
   el resultado debe ser idéntico y las pruebas deben pasar.
5. Entregar al usuario: `dist/FinancialTracker.gs` (para pegar en Apps Script), el `.patch` y los comandos de git.
   Los commits van en inglés.
6. **Si GitHub está atrasado** (el usuario aplicó un patch pero no hizo push), un patch contra GitHub choca con los
   cambios locales. En ese caso conviene entregar dos patches: uno solo con la versión nueva y otro acumulado, y
   pedir `git status` para elegir.

Comandos en **Windows PowerShell 5.1** (no acepta `&&`; `Copy-Item` necesita comas entre varios orígenes):

```powershell
cd "<carpeta del repo>"
git pull
git apply "..\agentic-fin-tracker-vX.Y.Z.patch"
git add -A
git commit -m "vX.Y.Z: <resumen en inglés>"
git push
```

## 6. Reglas de trabajo (las que han evitado o detectado errores)

**Calidad**
- **Reproducir antes de corregir.** Primero una prueba que falle con el error reportado; luego la corrección.
- **Pruebas de mutación** para cada protección nueva: copiar `src/` a una carpeta temporal, romper la línea que
  protege, correr `GS_ROOT=<carpeta> node --test tests/v1NN.test.js` y confirmar que la prueba falla. Si no falla,
  la prueba no protege nada y hay que mejorarla (varias veces el fixture no ejercitaba el caso real).
- **Cuando una prueba falla, verificar quién tiene razón.** Varias veces el error estaba en la prueba (un dato mal
  calculado, una fila mal ubicada) y el código era correcto. No ajustar el código para que pase una prueba equivocada.
- **Verificar con datos reales en memoria** cuando el usuario comparte un correo o PDF: pasarlo por el lector completo
  (`parseEmailMessage`), no solo por el extractor. Así se encontró que el pie de seguridad de QIK filtraba la compra.
- **Interfaces (diálogos, Wizard, panel):** además de las pruebas, ejercitarlas en un navegador real con Playwright
  (el HTML se obtiene del mock; `google.script.run` se simula) y revisar una captura.
- **Ser honesto con lo no verificable.** Lo que solo ocurre en Apps Script real (GOOGLEFINANCE, Drive, fechas pegadas
  desde Excel) se dice como hipótesis, se protege con verificaciones y registros, y se confirma con el usuario.

**Privacidad**
- **Nada personal en el repositorio** (es público): ni nombres, ni cuentas, ni tarjetas, ni montos, ni comercios de un
  usuario real. Los fixtures conservan la estructura del correo real con **datos inventados**.
- **Escanear antes de cada patch** con `grep` de los datos reales que se hayan visto en la conversación (nombres,
  números de cuenta, montos). Ese escaneo atrapó, por ejemplo, montos reales que habían quedado en comentarios.
- Nunca subir `.eml`, estados de cuenta, ni archivos `.tsv` con datos de un usuario.
- `tests/v124.test.js` acepta en el código solo remitentes de dominios de bancos; al agregar un banco, agregar su
  dominio a esa lista.

**Diseño**
- **Una prueba que fija un comportamiento viejo** ("como antes") se cambia solo a propósito, con evidencia (ejemplo:
  v1.1.63, un código de autorización "US34K7" ya no cuenta como dólares) y explicando el porqué en la prueba.
- **Nada específico de un usuario en el código.** Categorías y reglas personales van en **Custom Rules** o se
  categorizan a mano. El catálogo de tarjetas lleva solo **condiciones públicas del producto**: la tasa es el cashback
  base; puntos, millas y promociones por día o con vencimiento van en la nota. Nunca límites de crédito, fechas de
  corte o de pago, ni costos de un titular (esos van en el Setup Wizard de cada quien).
- **Las hojas derivadas** (Bank Transfers, Incoming Transfers, Raw_<BANCO>) **se reconstruyen desde Transactions** en
  cada actualización. Cualquier lista de hojas en el código debe incluir a todas (un olvido dejó Incoming Transfers
  sin formato). Mejor aún: derivar las listas de una sola fuente, como `BANK_ORDER`.
- **Una configuración vieja debe seguir funcionando** sin abrir el Wizard (ejemplo: el ingreso adicional único se migra
  a la lista nueva y el Dashboard lo muestra sin quedar en 0).

## 7. Trampas conocidas de Apps Script

- **Diálogos HTML:** el script de cada ventana vive dentro de un *template literal* de JavaScript. Un apóstrofo dentro
  de una cadena del script del cliente (`'account's'`) rompe todo el script de la ventana sin error visible. Usar
  textos sin apóstrofo o comillas dobles. `tests/v141.test.js` compila el script de cada diálogo y verifica que cada
  botón llame a una función existente.
- **GOOGLEFINANCE** tarda en cargar después de escribir la fórmula; mientras carga, la fórmula cae al último precio
  conocido. Holdings espera hasta que "Price source" diga *live* (`waitForLivePrices`).
- **Límite de 6 minutos por ejecución.** Las corridas largas se detienen a tiempo, guardan lo leído y dejan los pasos
  pesados para la siguiente (`⏸` y `⏭` en el resumen). El registro de ejecución muestra el tiempo de cada fase (`⏱`).
- **Gmail `getPlainBody()`** convierte el HTML del banco a texto a su manera; los extractores trabajan sobre el texto
  **aplanado** (espacios simples) para no depender de cómo reparte las celdas de una tabla.
- **Gmail escribe las negritas como `*texto*`** (`*COMERCIO: *UBER*EATS`, `*RD$ 20.00*`). Al parecer solo marca las
  etiquetas `<b>` y `<strong>`, no los estilos `font-weight:bold` (evidencia: el registro real de la plantilla
  "Servicio de Alerta - Nuevo Consumo" de LAFISE no trae asteriscos, y su HTML usa estilos para las etiquetas). Al
  emular desde un `.eml`, reemplazar cada `<b>`/`<strong>` por `'*' + texto + '*'` **pegado al texto**: insertar el
  asterisco como un fragmento aparte mete saltos de línea que Gmail no pone y produce falsos errores. Los extractores nuevos leen
  con `flatText()`, que quita esas marcas y conserva el asterisco dentro de una palabra. **Los fixtures deben llevar las
  marcas** donde el correo real tiene negritas: los de v1.1.60/61 no las tenían y las pruebas pasaban con correos que
  en la vida real fallaban. Para emular el texto de Gmail desde un `.eml`, convertir `<b>`, `<strong>` y
  `font-weight:bold` en `*…*`. La hoja Unrecognized muestra el texto tal como lo entrega Gmail: es la mejor referencia.
- **Fechas:** `new Date('2026-08-28')` es medianoche UTC, que en Santo Domingo es el día anterior. Un texto
  `yyyy-mm-dd` se toma como su día (`normalizeDateForCompare`).
- **Filas escritas a mano** en Bank Transfers e Incoming Transfers se importan a Transactions antes de reconstruir la
  hoja. Una vez, filas pegadas desde Excel quedaron sin fecha en una corrida que además murió por tiempo, y **no se
  pudo reproducir**. Quedó una verificación posterior al guardado que corrige la fecha y deja en el registro la línea
  "⚠️ … please report this line". Si aparece, investigar con esa evidencia.
- **No editar la hoja mientras corre el tracker** (pegar o borrar filas durante una ejecución).
- Para diagnosticar en la hoja real, una técnica útil: un archivo temporal `Diagnostico` en Apps Script con una función
  que llame a las funciones del tracker (que son globales) y registre resultados paso a paso; se borra al terminar.

## 8. Cómo agregar un banco o un formato nuevo

1. Pedir al usuario el **correo original**: en Gmail, ⋮ › *Show original* › *Download original* (`.eml`). Uno por
   cada tipo de alerta (consumo, transferencia enviada, recibida…).
2. Verlo con el lector completo, en memoria, sin guardar el contenido en el repo.
3. En `03_gmailMonitor.gs`: una entrada en `BANK_PATTERNS` (`name`, `fromDomain` —se compara con *contiene*—,
   `searchQuery`, `extractors`, `keywords`), el banco en `BANK_ORDER`, y si hace falta, el asunto en
   `TYPE_SUBJECT_KEYWORDS` (el asunto decide el tipo antes que el cuerpo).
4. Fixture en `tests/fixtures/` con la estructura real y datos inventados; pruebas con mutaciones; dominio del remitente
   en la lista de `tests/v124.test.js`.
5. Si tiene tarjetas, sus productos en `CARD_PRODUCTS` (solo condiciones públicas).

## 9. Pendientes y limitaciones conocidas

- **Estado de cuenta de Popular:** el lector de PDF ya lee su formato, pero falta un **correo original** con el estado
  adjunto para saber el remitente y el asunto. Mientras tanto, sus créditos y débitos se pueden pegar a mano en
  Incoming Transfers y Bank Transfers.
- **PDF protegidos con contraseña** (algunos bancos usan la cédula): hoy se detectan y aparecen en Unrecognized como
  *password-protected*. Implementar el descifrado necesita una muestra real. Si se hace, la cédula se guarda en las
  propiedades privadas del script, **nunca** en la hoja, los registros ni el repo.
- **Transferencia recibida de Banesco** ("Notificación de Transferencia Recibida"): escrita en v1.1.62 a partir de la
  enviada; falta el `.eml` para confirmar el campo del remitente.
- **LAFISE "Aviso de transferencia en banco local" exitoso:** solo hay muestras fallidas ("Estado: Error"); cualquier
  estado que no sea un fallo se lee como transferencia. Con un `.eml` exitoso, confirmar la palabra del estado.
- **Resumen diario, "% sobre tu promedio":** compara un día contra el promedio de 30 días (sin facturas ni costos
  fijos), y un gasto grande pero esporádico (combustible, una compra grande) dispara porcentajes enormes. Propuesta en
  discusión: comparar contra un "día típico" robusto (mediana o percentil) y nombrar los gastos esporádicos aparte.
- **Uber en LAFISE:** aparecen retenciones "UBR* PENDING.UBER.COM" con el mismo monto que un "UBER*RIDES" minutos
  después. Si se confirma que son retenciones, una Custom Rule a Exclude evita contarlas dos veces.
- **Sin muestra todavía:** el correo de bienvenida de BDI (hoy se filtra por palabras: "bienvenido/a", "primer
  depósito", "apertura de cuenta"; con su `.eml` se puede confirmar), consumos de Scotiabank en DOP, consumos con
  tarjeta de débito de Banreservas, otros recibos de la app de Banreservas (pagos de servicio, entre cuentas propias),
  correos de BHD con varias transacciones agrupadas, acreditación de cashback.
- **Cashback por categorías** (topes, mínimos, días): el catálogo solo guarda el cashback base y una nota.
- **El formato del Dashboard** tarda ~80 s en hojas grandes; vigilar el límite de 6 minutos si crecen los datos.
- **IBKR:** pendiente leer su reporte (Flex Query) cuando haya operaciones.

- **Consumos de QIK que no aparecían:** resuelto en v1.1.62 (las negritas). **Consumos de Scotiabank (v1.1.61):** el correo real se lee bien en memoria; la hipótesis es que el
  banco está sin marcar en la configuración. Desde v1.1.61 el resumen nombra los bancos sin marcar que tienen correos y
  cuenta lo leído por banco. Confirmar con el usuario.
- **Transferencias recibidas en Banreservas y Scotiabank:** el correo no nombra al titular, así que no se sabe si es
  dinero propio. Se resuelve con una Custom Rule (nombre propio → Exclude).

## 10. Dos usuarios, un repositorio

El código es el mismo para ambos: lo que cambie uno sirve al otro. Para no divergir:
- Todo cambio parte de la **última versión de GitHub** (`git pull` antes de empezar).
- Cada versión actualiza **este archivo** (estado, bancos, trampas, pendientes) además del CHANGELOG. Así, el
  siguiente Claude, en Claude Code o en el chat, parte del contexto más reciente sin pasarse archivos a mano.
- Nada personal en este archivo: es público, como el resto del repositorio. Lo de cada uno va en `CLAUDE.local.md`.
