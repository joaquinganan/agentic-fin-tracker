# Guía de uso — Financial Tracker

El tracker lee las alertas de tus bancos en Gmail, registra cada consumo y transferencia en una hoja de Google, las
categoriza y te envía un resumen diario y otro mensual. También puede llevar tus inversiones: las órdenes y dividendos
de HAPI llegan solos desde sus correos, y los fondos o la pensión se registran con su saldo.

Todo vive en **tu propia hoja y tu propia cuenta de Google**. Nadie más ve tus datos. La interfaz (menús, hojas y
paneles) está en inglés; esta guía usa los nombres tal como aparecen en pantalla.

> ¿Algo no está claro o no funciona? Abre **📊 Tracker › 📘 Start here**: te dice qué falta y qué hacer.

---

## 1. Instalación (5 minutos)

1. Crea una **hoja de cálculo nueva** en Google Sheets (por ejemplo, "Financial Tracker").
2. En la hoja, abre **Extensions › Apps Script**. Tiene que ser desde la hoja: el tracker vive "pegado" a ella.
3. Abre el archivo **[dist/FinancialTracker.gs](../dist/FinancialTracker.gs)** del repositorio, copia **todo** su
   contenido y pégalo en `Code.gs`, reemplazando lo que haya. Ese único archivo es el tracker completo.
4. Guarda (**Ctrl+S**) y **recarga la hoja**. Aparece el menú **📊 Tracker**.
5. La primera vez que uses una opción del menú, Google te pedirá permisos: leer tu Gmail (para las alertas del banco),
   editar la hoja y enviarte los correos de resumen. Acéptalos con tu cuenta.

## 2. Configuración inicial

Abre **📊 Tracker › 🔧 Setup Wizard**. Te pide:

- **Tu correo** y los **bancos** cuyas alertas llegan a tu Gmail.
- **Tus ingresos** (sueldo y otros ingresos, en pesos o dólares) y cómo calcular los descuentos (ARS, AFP e ISR, de
  forma automática con las reglas de nómina dominicanas o con tus montos). Con eso el tracker compara tu gasto contra lo
  que realmente te queda.
- Tus **tarjetas de crédito**, con su día de corte y de pago.
- Los **correos de resumen** que quieres recibir: diario, mensual o ambos, y a qué hora.

Al guardar, el tracker crea las hojas y programa la **actualización diaria a las 6 AM**. Termina con "Next: 📘 Start
here".

## 3. 📘 Start here: tu lista de pendientes

**📊 Tracker › 📘 Start here** abre un panel al lado de la hoja con cada paso: configuración, actualización diaria,
correos del año, correos que no se pudieron leer, transferencias sin categoría y, si llevas inversiones, sus pendientes.

Los pasos **se revisan solos contra tu hoja**: no se marcan a mano. Por eso también sirve para detectar problemas: si
algo se rompe después (una corrida que dejó de programarse, un correo que no se pudo leer), vuelve a aparecer con un botón
para resolverlo. Ábrelo al terminar el Wizard y cada vez que algo te parezca raro.

## 4. Primera lectura: los correos del año

Para cargar lo que ya pasó, usa **📊 Tracker › 📅 Monitor by Date Range** del 1 de enero a hoy.

Google corta cualquier ejecución a los **6 minutos**. Si tienes muchos correos, el tracker se detiene antes, guarda lo que
leyó y el resumen lo dice ("⏸ Stopped reading early…" o "⏭ Left for the next run…"). **Repite la misma búsqueda** hasta
que esas líneas no aparezcan: los correos ya leídos se saltan, así que cada vuelta llega más lejos.

## 5. El día a día

**Lo que pasa solo:**

- Cada mañana a las 6 AM se leen las alertas nuevas y se actualizan las hojas y el Dashboard.
- Si los activaste, recibes el **resumen diario** (lo que gastaste ayer, cómo vas en el mes, recomendaciones) y el
  **resumen mensual** el día 1.
- **📊 Tracker › 🔄 Monitor Gmail Now** hace la misma actualización en el momento.

**Lo que conviene revisar de vez en cuando:**

- **Transferencias sin categoría.** Aparecen en naranja en **Bank Transfers**. Escribe la categoría en esa hoja o en
  **Transactions**; usa **Exclude** para movimientos entre tus propias cuentas.
- **Unrecognized.** Los correos del banco que el tracker no pudo leer, con el motivo, un fragmento del correo y un enlace
  **Open** para abrirlo en Gmail. Si no te importa, pon su *Status* en **Ignore**. Si es un formato nuevo, ver la sección
  10.

## 6. Categorías

- Cada consumo recibe una categoría automática según el nombre del comercio.
- **Puedes cambiarla a mano** en **Transactions** o en **Bank Transfers**: el tracker respeta lo que pusiste y no lo
  vuelve a cambiar. Para volver a la categoría automática, **borra la celda**.
- Para que algo se categorice siempre igual (un comercio, una persona a la que le transfieres), agrega una regla en
  **Custom Rules**: tu correo, la categoría y una palabra que aparezca en el comercio o beneficiario. Puedes crear
  categorías nuevas escribiendo su nombre.
- **📊 Tracker › 🔁 Recategorize Saved Transactions** vuelve a aplicar las reglas a todo lo guardado.
- Las hojas **Raw_** (una por banco) son copias: lo que edites ahí no se guarda.
- Si borras una fila de **Transactions**, la siguiente corrida que cubra esa fecha la vuelve a traer desde el correo.

## 6b. Transferencias entrantes: lo que te devuelven

Si compartes gastos (la renta o los servicios con un roommate, por ejemplo) y la otra persona te transfiere su parte,
el tracker puede **restarla de lo que gastaste** en esa categoría.

- Lo recibido aparece en la hoja **Incoming Transfers**, sin categoría y resaltado en naranja.
- **Dale la categoría que devuelve.** Si pagaste 30,000 de renta y te transfirieron 15,000 con categoría *Rent*, la
  renta muestra 15,000 en el Dashboard y en los resúmenes. Usa **Exclude** para dinero que no devuelve ningún gasto.
- Si la transferencia trae el nombre de quien envía (LAFISE lo trae), una regla en **Custom Rules** con ese nombre y la
  categoría lo hace sola desde entonces.
- Lo que te transfieres **desde tu propia cuenta** se reconoce y queda en *Exclude*.
- La categoría que escribes en Incoming Transfers (o en Bank Transfers) se aplica **al instante**: el Dashboard cambia
  sin esperar a la siguiente actualización. Recuerda que cada transferencia cuenta en el mes de su fecha: lo que llegó
  en agosto se resta de agosto.
- **¿Una transferencia que el tracker no vio?** Escríbela tú en una fila vacía de **Incoming Transfers**: fecha
  (yyyy-mm-dd), banco (opcional), de quién, categoría y monto (positivo; moneda DOP si la dejas vacía). En la siguiente
  actualización, o con **🔁 Recategorize Saved Transactions**, se guarda y queda como las demás. Si le falta la fecha o
  el monto, no se pierde: aparece en Unrecognized con lo que escribiste.

**Banesco** no avisa la mayoría de las transferencias que recibes, pero su **estado de cuenta mensual** (el PDF que llega
por correo) las lista. El tracker toma **todos los créditos** del estado, con su descripción tal cual ("Ach Ibanking",
"Lbtr <nombre>", depósitos, intereses…): los categorizas a mano o con una Custom Rule sobre su descripción, y *Exclude*
para lo que no devuelve ningún gasto. Los que nombran al titular de la cuenta se reconocen como tuyos. El tracker lee el
PDF por sí solo, sin activar nada, y comprueba que lo leído cuadre con los totales del propio estado de cuenta; si no
cuadra, no toma nada y lo avisa en Unrecognized.

## 7. Inversiones

### HAPI: primeros pasos

1. **Tus posiciones actuales.** En la web de HAPI abre el portafolio y selecciona desde **Total balance** hasta el último
   activo (baja hasta el final de la lista). Copia, y en **📊 Tracker › 📋 Paste Broker Positions** pega, revisa la
   vista previa y guarda. Si incluyes *Total assets*, el tracker comprueba que no falte ninguna posición; si falta
   algo, te dice cuánto. *Total money* se registra como el efectivo.
2. **El valor al inicio del año**, para medir tu rendimiento anual. Búscalo en el estado de cuenta de diciembre de HAPI:
   es el **valor de cierre** ("closing balance", *no* el de apertura). Regístralo en **📊 Tracker › ➕ Add Balance or
   Deposit**: cuenta *HAPI*, tipo *Broker*, fecha del cierre, *A balance*, en USD.
3. **Corre el Monitor by Date Range** del año, si no lo has hecho: trae tus compras, ventas y dividendos desde los correos
   de HAPI.

### Qué se registra solo

- **Compras y ventas** (con su comisión) y **dividendos**, desde los correos de HAPI.
- **Depósitos**, desde tus transferencias bancarias a la cuenta de HAPI (en República Dominicana, "OUROSR SRL"). La
  palabra que las identifica está en la hoja **Investment Accounts**.
- HAPI también envía un correo "Deposit Completed" **sin el monto**. El tracker lo guarda como aviso y, si no encuentra el
  depósito correspondiente, lo lista en **Unrecognized** como *Deposit without amount*. Agrega el monto con **➕ Add
  Balance or Deposit › A deposit** (con la fecha de HAPI) y la fila desaparece sola.
- Si HAPI no envió el correo de algún dividendo, puedes agregarlo como una fila *Dividend* en el **Investment Ledger**.

### Fondos y pensión

Con cada estado de cuenta, usa **➕ Add Balance or Deposit**: el fondo con *Units × unit price* (cuotas y valor de la
cuota) y la pensión con *A balance*. Cada estado de cuenta es una **fila nueva**: no edites las anteriores, que son el
historial. Si guardas otra vez la misma cuenta con la misma fecha, se actualiza en lugar de duplicarse. Para medir el
rendimiento del año, registra también el saldo al inicio del año.

### Cómo leer Holdings

- **Positions:** cada acción con su precio actual (acciones y ETFs desde Google Finance; cripto desde Coinbase), valor,
  ganancia no realizada y peso.
- **Performance:** la ganancia de cada cuenta es *valor actual − valor inicial − dinero que metiste*. Las compras no
  cuentan como aporte, porque solo convierten efectivo en acciones; los **depósitos** sí. El rendimiento anualizado
  aparece a partir de 180 días.
- Si lo que compraste supera lo que depositaste, Performance lo advierte: probablemente falta registrar algún depósito, y
  ese dinero se estaría contando como ganancia.
- **📊 Tracker › 📈 Refresh Investments** actualiza todo al momento.

## 8. Actualizar a una versión nueva

1. En **Extensions › Apps Script** abre `FinancialTracker.gs`.
2. Selecciona todo (**Ctrl+A**), pega la versión nueva y guarda (**Ctrl+S**).
3. Recarga la hoja. La segunda línea del archivo dice la versión, y **📘 Start here** también la muestra.

Tus datos, configuración y corridas programadas no se tocan.

## 9. Problemas comunes

| Qué ves | Qué hacer |
|---|---|
| **No aparece el menú 📊 Tracker** | Casi siempre es un pegado incompleto o duplicado. En Apps Script, elige la función `onOpen` y pulsa **Run**: el error dice qué pasa. Si dice *"has already been declared"*, el código está pegado dos veces: deja un solo `FinancialTracker.gs`. |
| **Un aviso se queda abierto** | La ejecución terminó o se cortó. Recarga la hoja. |
| **"Exceeded maximum execution time"** | Repite la misma búsqueda: cada vuelta avanza. |
| **Una fila en rojo en Transactions** | El tracker no pudo leer el comercio de ese correo. Aparece también en Unrecognized. |
| **Una fila en rojo en el Investment Ledger** | Le falta la fecha y no se está contando. Escríbela en la columna A. |
| **Un saldo o una cifra no cuadra** | Abre **📘 Start here** y revisa sus advertencias. |

Otras opciones del menú:

- **📊 Tracker › 📊 Dashboard** abre el resumen principal.
- **📊 Tracker › 📬 Send Daily Summary Now** y **🗓️ Send Monthly Summary Now** envían los resúmenes en el momento, para
  probarlos.
- **📊 Tracker › 🙈 Show / Hide Settings Tabs** oculta o muestra las hojas de configuración (Investment Accounts,
  Configuration, Categories).
- **📊 Tracker › ⚙️ View Config** muestra tu configuración actual.
- **📊 Tracker › 🗑️ Reset System** borra **todo lo que el tracker creó**: la configuración, Transactions, Bank Transfers,
  el Dashboard, las hojas Raw_, Custom Rules, Unrecognized y las corridas programadas. **No se puede deshacer.** Solo
  conserva las hojas de inversiones, porque tienen lo que escribiste a mano. Después, el Setup Wizard y un Monitor by
  Date Range reconstruyen lo demás desde tus correos.

## 10. Agregar un banco o una tarjeta nueva

Los correos de un banco que el tracker aún no sabe leer aparecen en **Unrecognized**. Para agregarlo:

1. Abre el correo con el enlace **Open**.
2. En Gmail: **⋮ › Show original › Download original**. Eso descarga un archivo `.eml`.
3. Envíalo a quien mantiene el tracker, con un ejemplo de **cada tipo** de alerta de ese banco (consumo, transferencia,
   retiro…). Para el cashback de una tarjeta, agrega el **reglamento oficial** del programa.

Los `.eml` tienen tus datos: compártelos solo en privado, **nunca** en el repositorio público.

## 11. Privacidad

- Tus datos quedan en tu hoja y tu cuenta de Google.
- El tracker solo sale de Google para consultar **precios públicos de criptomonedas** en Coinbase, sin enviar ningún
  dato tuyo.
- El código del repositorio no contiene datos personales de nadie.
