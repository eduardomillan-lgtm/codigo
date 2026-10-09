# Motor de reclutamiento de agentes · KW Market Center Marbella

Sistema para localizar agentes inmobiliarios en la Costa del Sol occidental,
puntuarlos por probabilidad de incorporación y llevarlos por una secuencia de
seguimiento hasta la entrevista con la Team Leader.

Se monta sobre el mismo Google Apps Script que ya usas para el Dashboard, así
que no hay que instalar nada nuevo ni pagar otra herramienta.

---

## 1. Instalación (5 minutos)

| Paso | Acción |
|---|---|
| 1 | En el proyecto de Apps Script, crea un archivo nuevo y pega **`reclutamiento.gs`**. |
| 2 | Crea un archivo HTML llamado **`panel_reclutamiento`** y pega **`panel_reclutamiento.html`**. |
| 3 | En tu `onOpen()` del archivo `gs` (línea ~198), añade antes del `.addToUi()` final una línea nueva: `recCrearMenu();` |
| 4 | Ejecuta **`recInicializarTodo()`** desde el editor. Crea las 8 hojas y siembra los Smart Plans. |
| 5 | Rellena la hoja **`Rec_Config`**: nombre de la TL, teléfono, email y dirección del MC. |
| 6 | Menú **🎯 Reclutamiento → Configurar claves de API**. Mínimo: `GEMINI_API_KEY` y `PLACES_API_KEY`. Para el X-ray, además `CSE_API_KEY` y `CSE_CX`. |
| 7 | Menú **🎯 Reclutamiento → Activar automatización diaria**. |

> ⚠️ `reclutamiento.gs` **no** define `onOpen()` a propósito: ya tienes dos en
> el archivo `gs` (líneas 54 y 198) y el segundo gana. Añadir un tercero
> rompería el menú.

---

## 2. ¿De dónde salen los contactos?

Esta es la pregunta de verdad. Ordenadas por **rentabilidad real**, no por lo
llamativas que suenen:

| # | Fuente | Qué te da | Volumen estimado | Automatizable | Coste |
|---|---|---|---|---|---|
| 1 | **Webs de las agencias** (páginas de equipo) | Agente por agente: nombre, cargo, email, teléfono, idiomas | **1.500–4.000** | ✅ Total | Céntimos de Gemini |
| 2 | **Google Places API** | El censo de agencias: nombre, web, teléfono, dirección, reputación | **300–600 agencias** | ✅ Total | ~15–25 € el barrido completo |
| 3 | **X-ray de LinkedIn vía Google** | Nombre, cargo, agencia y URL del perfil. **Sin teléfono** | 500–2.000 perfiles | ✅ Total | 100 consultas/día gratis |
| 4 | **Portales** (Idealista, Fotocasa, Kyero, Resales) | **Producción real**: nº de inmuebles publicados y rango de precio | El mismo universo, pero cualificado | ⚠️ Semi (copiar/pegar) | 0 € |
| 5 | **LinkedIn Sales Navigator** | Lo del X-ray más antigüedad, **alertas de cambio de trabajo** y exportación | 2.000–5.000 perfiles | ⚠️ Semi (export manual) | ~100 €/mes |
| 6 | **Referidos de tus propios agentes** | Pocos, pero con la conversión más alta de todas | 5–15/mes con incentivo | ❌ Proceso humano | Incentivo interno |
| 7 | **Formación abierta en la oficina** | Candidatos que vienen ellos | 10–30 por evento | ❌ | Organización |
| 8 | **Instagram** | Agentes con marca personal y cierres visibles | Cientos | ❌ Manual | 0 € |
| 9 | **Ofertas de empleo de la competencia** (InfoJobs, Indeed) | Qué agencias tienen rotación → dónde hay gente descontenta | Señal, no contactos | ⚠️ Semi | 0 € |
| 10 | **Registro de Agentes Inmobiliarios de Andalucía** | El censo oficial completo | Todo el sector residencial | 🔜 Cuando abra | Gratis |

> **Si no tienes Sales Navigator** (fuente 5), la 3 lo cubre casi entero y es
> gratis. La diferencia real es que Sales Navigator te avisa cuando alguien
> **cambia de trabajo**, que es la señal de mayor conversión que existe. Si no
> lo tienes, relanza el X-ray cada mes y compara contra la base: los perfiles
> nuevos suelen ser movimientos recientes.

### El flujo que de verdad funciona

```
Google Places  →  censo de agencias de la zona          (automático)
      ↓
Webs de agencias  →  agentes con nombre y contacto      (automático)
X-ray de Google  →  perfiles que no salen en esas webs  (automático)
      ↓
Scoring  →  los 200 mejores                             (automático)
      ↓
Portales (pegado)  →  cuántos inmuebles tiene cada uno  (semi: 2 h/semana)
      ↓
LinkedIn  →  antigüedad y señal de cambio reciente      (semi)
      ↓
Smart Plan  →  llamada, WhatsApp, LinkedIn, email       (automático)
      ↓
Entrevista con la Team Leader
```

**Los pasos automáticos te dan la base entera en una tarde.**
Los pasos semiautomáticos son los que cualifican, y son los que de verdad
mueven la aguja: un agente con 30 inmuebles en Nueva Andalucía vale veinte
veces lo que un nombre suelto en una web.

### Sobre el Registro de Andalucía

La **Ley 5/2025 de Vivienda de Andalucía** (BOJA 24/12/2025, en vigor desde el
**24 de enero de 2026**) creó el *Registro de Agentes Inmobiliarios
Especializados del Sector Residencial*: **público, gratuito y obligatorio**.
Sin inscripción no se puede intermediar en vivienda residencial en Andalucía.

Cuando esté operativo será **la mejor fuente posible**: el censo íntegro, legal
y actualizado. Pero **a día de hoy el registro todavía no está en marcha** — la
Junta tiene de plazo hasta el **24 de enero de 2028** para desarrollar el
reglamento y ponerlo a funcionar.

**Acción:** revisa el BOJA cada trimestre. El día que abra, es la primera
fuente a la que ir. Te ahorra todo el paso 1 y 2.

### Para los que "no están en agencia"

Es tu segmento de mayor conversión y casi nadie lo trabaja:

- **Autónomos inmobiliarios.** Andalucía ha sumado ~12.500 autónomos del sector
  en cinco años. Ya venden solos, sin marca, sin estructura y sin nadie a quien
  preguntar. La propuesta de KW les encaja mejor que a nadie. Búsqueda nº 2 del
  generador de LinkedIn.
- **Cambio de sector.** Hostelería de lujo, banca privada, retail premium,
  náutica, clubes de golf. Ya tratan a diario con el cliente que compra una
  villa de 2 M€ en Marbella. Eso no se enseña; lo técnico sí. Búsqueda nº 5.
- **Idiomas de los mercados compradores.** Sueco, neerlandés, alemán, ruso,
  árabe, polaco. Un agente que habla el idioma del comprador accede a cartera
  que el resto no puede atender. Búsqueda nº 6.

---

## 3. La secuencia de canal (esto no es estético)

```
1. LLAMADA        ← al teléfono profesional que el agente publica
2. WhatsApp       ← solo si no contesta: una línea, identificado, con salida
3. WhatsApp       ← libre, una vez hay conversación
4. Email          ← para lo que no cabe en un WhatsApp, con la info del art. 14
```

**Nunca al revés, y hay dos razones.**

**La de plataforma.** La WhatsApp Business Platform exige consentimiento previo
antes de abrir conversación con plantillas. Un envío masivo en frío te degrada
la calificación del número, te reduce el límite de mensajes y acaba en bloqueo.
Desde octubre de 2025 Meta además limita los mensajes a contactos desconocidos
que nunca responden.

**La legal.** El art. 21 LSSI-CE exige consentimiento previo para
comunicaciones comerciales por vía electrónica. La AEPD ya ha sancionado
exactamente este supuesto: multa a una empresa por guardar los datos de una
persona cuyo CV era público en LinkedIn y enviarle un email comercial sin
consentimiento, y otra por enviar comunicaciones electrónicas a un profesional
sin su autorización.

La llamada a un teléfono que el propio agente publica **para recibir llamadas
profesionales** es, con diferencia, el primer contacto más defendible. Y además
convierte mucho más.

> Si tienes **Kelly** (la capa de WhatsApp de KWSync), los toques 3 en adelante
> pueden ir por ahí una vez el agente ha consentido. El primero, no. Sección 9.

### El opt-in se consigue en la propia llamada

Al final del guion de apertura hay una frase que parece de relleno y no lo es:

> *"Te mando por WhatsApp el desglose de la zona que te comentaba. ¿Este número es el bueno?"*

Ese **"¿este número es el bueno?"** es tu consentimiento de WhatsApp. En cuanto
dice sí, el sistema lo marca (`Consentimiento_WA = SI`) y desde ahí el canal
queda abierto. El panel lo marca solo cuando pulsas *Interesado* o
*Entrevista agendada*.

---

## 4. Los Smart Plans

Cinco planes sembrados en la hoja `Rec_SmartPlan`. **Edítalos ahí**, no en el
código: lo que cambies se usa en la siguiente generación de cola.

### `AGENTE_ACTIVO` — 15 toques en 90 días

| Día | Canal | Qué se hace |
|---|---|---|
| −1 | — | Investigación. Sin un dato concreto suyo, no se llama |
| 0 | 📞 | Apertura. **No se vende KW**: se piden dos opiniones de mercado |
| 0 | 💬 | Solo si no contesta. Una línea, identificado, con salida |
| 2 | 💬 | Dato de mercado de **su** zona. Cero KW en el mensaje |
| 5 | 📞 | Segundo intento, en otra franja horaria |
| 7 | in | Conexión de LinkedIn con nota |
| 10 | 💬 | Caso real de un agente comparable |
| 14 | 💬 | Invitación a formación abierta en la oficina |
| 21 | 📞 | Check-in apoyado en lo que ya le has dado |
| 28 | 💬 | Calculadora de ingresos: la rellena él, no te da datos |
| 35 | ✉️ | Modelo económico en PDF + información del art. 14 RGPD |
| 45 | 💬 | Prueba social: alguien que acaba de incorporarse |
| 60 | 📞 | Reapertura con una razón nueva |
| 75 | 💬 | Informe trimestral de su zona |
| 90 | 📞 | Cierre de ciclo: o nurture mensual, o se archiva |

### Los otros cuatro

- **`AUTONOMO`** — su dolor no es el split, es la soledad y la falta de
  estructura. El guion ataca eso.
- **`CAMBIO_SECTOR`** — no saben que son candidatos. Hay que explicárselo, y
  luego resolver el miedo real: *"¿de qué vivo mientras aprendo?"*.
- **`NURTURE`** — un dato útil al mes, cero presión. **Una parte importante de
  las incorporaciones sale de aquí**, no de la primera conversación.
- **`POST_ENTREVISTA`** — incluye la estructura de **Career Visioning**: seis
  preguntas sobre su vida, no sobre el puesto. Tú preguntas y te callas; él se
  vende a sí mismo el cambio.

### Los corchetes son deliberados

Los mensajes llevan huecos `[así]`. **Rellénalos con datos reales de tu Market
Center.** Si no puedes respaldar una cifra, bórrala. Un dato inventado delante
de un agente que conoce la zona mejor que tú te cierra la puerta para siempre,
y encima corre entre la competencia.

### Activos que los mensajes prometen

| Paso | Activo | Estado |
|---|---|---|
| Día 2 y 75 | Desglose de mercado por urbanización | Lo tienes que sacar de tu CRM |
| Día 10 | Caso real de un agente comparable | Pídele permiso a la persona |
| Día 14 | Formación o mesa redonda abierta | Organizativo |
| **Día 28** | **Calculadora de ingresos** | ✅ **`calculadora_ingresos_agente.xlsx`** |
| Día 35 | Modelo económico en PDF | Material corporativo del MC |
| Día 45 | Agente recién incorporado dispuesto a hablar | Pídeselo |

---

## 5. Rutina de la Team Leader

### Cada día (20–30 min)
1. Abre **🎯 Reclutamiento → Panel diario**.
2. Trabaja la cola de arriba abajo: ya viene ordenada por prioridad.
3. Botón verde de WhatsApp → se abre con el texto escrito. Lo revisas y envías.
4. Marca el resultado. **Esto es lo único obligatorio**: sin ello el sistema no
   avanza el pipeline.

### Cada semana (1–2 h)
- Pega en el sistema las fichas de portal de los 20 candidatos mejor puntuados
  (**Captar candidatos → 4**). Es lo que convierte un nombre en un candidato.
- Revisa los que llevan 3 toques sin respuesta: ¿el guion o el candidato?
- Pide referidos en la reunión de equipo. Fuente nº 5, la de mejor conversión.

### Cada mes
- Lanza otro lote de **Extraer agentes de las webs** (van de 25 en 25 por el
  límite de 6 minutos de Apps Script).
- **Exportar para CommandMC** con los que ya están en conversación.
- Revisa el embudo: ¿dónde se cae la gente?

### Números de referencia para calibrar

Estimaciones del sector para que empieces a medir, **no verdades**. Sustitúyelas
por tus propios ratios en cuanto tengas tres meses de datos:

```
1 incorporación   ←  3–5 entrevistas
1 entrevista      ←  8–12 conversaciones reales
1 conversación    ←  4–6 intentos de contacto
─────────────────────────────────────────────
1 incorporación   ≈  150–350 toques
```

Para **2 incorporaciones al mes** necesitas una base activa de 500–800
candidatos y unos 40 toques diarios, que es el tope que trae configurado
`MAX_TOQUES_DIA`.

---

## 6. Qué es automatizable y qué no

**Honestamente**, porque aquí es donde fallan estos proyectos:

| ✅ Automático de verdad | ⚠️ Semiautomático | ❌ Es trabajo humano |
|---|---|---|
| Censo de agencias | Cualificar con portales | La llamada |
| Extraer agentes de webs | Exportar de LinkedIn | La entrevista |
| Puntuar y ordenar | Rellenar los corchetes | Career Visioning |
| Generar la cola diaria | Revisar el texto antes de enviar | Pedir referidos |
| Redactar el mensaje | | Cerrar la incorporación |
| Deduplicar y purgar | | |
| Bloquear las bajas | | |

Lo que este sistema te quita es **buscar, ordenar, redactar y acordarse**. Lo
que no te quita, y no debe, es hablar con la gente.

### Lo que NO debes hacer

- **Scrapers de LinkedIn** (Phantombuster, Apify y similares). Incumplen las
  condiciones de LinkedIn y la AEPD ya ha sancionado el uso de datos de perfiles
  públicos para contacto no consentido. El riesgo no compensa.
- **Rastrear portales automáticamente.** Idealista, Fotocasa y el resto lo
  prohíben en sus condiciones y en `robots.txt`. Por eso ese paso es de copiar
  y pegar: así es una persona consultando una web pública, que es para lo que
  está publicada.
- **Comprar bases de datos de agentes.** No tienen base legal trazable. El
  responsable del tratamiento acabas siendo tú.
- **Envíos masivos en frío por WhatsApp.** Pierdes el número y te expones.

---

## 7. RGPD y LSSI: antes de la primera campaña

El sistema trae la capa de cumplimiento montada:

- **Base legal por ficha** (interés legítimo, art. 6.1.f RGPD) y **origen
  registrado** en `Fuente` y `URL_Fuente`. Si alguien pregunta de dónde has
  sacado su teléfono, hay respuesta.
- **Ponderación de interés legítimo** redactada en la hoja `Rec_RGPD`.
- **Información del art. 14 RGPD** incorporada al email del día 35.
- **Lista de supresión** que bloquea **todos** los canales de golpe.
- **Salida en cada mensaje escrito** (*"dime baja y no te vuelvo a escribir"*).
- **Purga por retención** a los 365 días sin actividad.
- **Auditoría completa** de cada contacto en `Rec_Toques`.

### Las cuatro cosas que tienes que hacer tú

1. **Rellenar `Rec_RGPD`** con la razón social, NIF y dirección del MC.
2. **Que tu asesor de protección de datos valide la ponderación.** Lo que hay
   escrito es un punto de partida sólido, **no un dictamen jurídico**.
3. **Dar de alta esta actividad** en tu Registro de Actividades de Tratamiento
   (art. 30 RGPD).
4. **Restringir el acceso a la hoja** a la TL y a dirección. Hoy es una hoja de
   cálculo con datos personales de miles de personas.

### El riesgo concreto, dicho claro

Existe un debate real sobre si una aproximación de reclutamiento es
"comunicación comercial" a efectos del art. 21 LSSI. El art. 19 LOPDGDD
presume el interés legítimo en datos de contacto profesional, pero lo hace
pensando en mantener relación **con la empresa**, no en captar a la persona
para que se vaya de ella.

Por eso este sistema **arranca por teléfono** y deja el email y el WhatsApp
para después del consentimiento. **No inviertas ese orden sin asesoramiento.**

---

## 8. Encaje con CommandMC

CommandMC ya tiene módulo de *Recruits* con Recruit Management, Smartviews y
SmartPlans de email, SMS y tareas. **Lo que no tiene es WhatsApp**, y los SMS
automáticos requieren una cuenta de Twilio conectada.

Reparto recomendado:

| | Este sistema | CommandMC |
|---|---|---|
| Captación y fuentes | ✅ | ❌ |
| Scoring y priorización | ✅ | Parcial |
| WhatsApp y llamada | ✅ | ❌ |
| Expediente oficial del recruit | ❌ | ✅ |
| Email y tareas del equipo | Parcial | ✅ |
| Reporting a la región | ❌ | ✅ |

**Flujo:** captas y calientas aquí; en cuanto hay conversación real,
**Exportar para CommandMC** y el expediente vive allí. Así no duplicas trabajo
ni pierdes el reporting oficial.

---

## 9. Encaje con Kelly (KWSync)

Kelly ya es la capa de WhatsApp de KWSync: Cloud API, plantillas aprobadas por
Meta, página de consentimiento en siete idiomas con prueba documental, gestión
de BAJA, ventana de 24 h, seguimientos, reactivación, campañas y topes de envío.
Todo eso está resuelto y **no hay que construirlo otra vez**.

Pero el reclutamiento **no puede ir por el mismo camino que los leads de
cliente**, por tres razones:

**1. El número.** El de Kelly es el de la oficina, y es el que contesta a los
leads en menos de un minuto. Ese número es infraestructura de ingresos. La
prospección en frío a agentes de la competencia es justo el tráfico que genera
bloqueos y denuncias, y un número con la calificación caída pierde límite de
envío. No se arriesga la respuesta a leads para ahorrarle clics a la TL.

**2. La base legal.** La premisa de Kelly es *"contestar a lo que el cliente ha
preguntado no necesita más permiso: la consulta la hizo él"*. Un agente al que
reclutas **no ha preguntado nada**. No hay consulta previa, así que esa premisa
no te cubre.

**3. Las plantillas.** Las de Kelly son de servicio. Una de reclutamiento es
categoría **MARKETING** para Meta: otras reglas, otro precio y más rechazos en
revisión. Y reutilizar una plantilla de servicio para prospección es motivo de
sanción, porque la autorización va ligada a la finalidad.

### El reparto que sí funciona

| Toque | Canal | Por qué |
|---|---|---|
| 1 · llamada | Teléfono de la TL | El canal más defendible y el que más convierte |
| 2 · WhatsApp si no contesta | `wa.me` desde el móvil de la TL | Riesgo cero para el número de la oficina |
| **3 en adelante, con consentimiento** | **Kelly** | Ya tiene plantillas, consentimiento y topes |

Para activarlo: `MODO_WHATSAPP = KELLY` en `Rec_Config`. Los toques de WhatsApp
de candidatos **con consentimiento** dejan de generar enlaces y se escriben en
`Rec_Cola_Kelly` con teléfono, idioma, plantilla sugerida y variables, listos
para que Kelly los consuma. Sin consentimiento no se entrega nada: se reencamina
a LinkedIn o a llamada, igual que antes.

**Decide el número antes de pedir plantillas.** Lo razonable es un **segundo
número**, a nombre del Market Center y no de la oficina, solo para
reclutamiento: separa reputaciones y separa facturación. Menú → Kelly/KWSync →
*Ver plantillas de Meta que faltan* te saca las cuatro que harían falta.

### Kelly también cualifica, y ahí está el valor de verdad

Kelly no solo envía: **conversa y cualifica**. Lo que hace con un comprador
—una pregunta por mensaje, con botones, sin repetir, parando cuando toca y
avisando si piden que les llamen— sirve igual para un candidato.

Y resuelve el punto más débil del sistema. El score se calcula con producción
**estimada** de los portales, y eso infravalora a un perfil concreto: el agente
que trabaja por referidos y publica poco. Caso real del sistema:

| Fuente del dato | Producción | Score | Prioridad |
|---|---|---|---|
| Estimada del portal | 3 anuncios de ~280.000 € | 58 | B · llamar este mes |
| **Dicha por él a Kelly** | **23 operaciones, honorarios 20-40k** | **76** | **A · llamar esta semana** |

El mismo agente. Kelly lo rescata de la cola larga.

**Cinco preguntas, no diez:**

1. ¿Cuántas operaciones cerraste el año pasado? *(botones)*
2. ¿Honorarios medios por operación? *(botones)*
3. Si pudieras cambiar UNA cosa de cómo trabajas hoy, ¿cuál sería? *(abierta)*
4. ¿Qué te haría considerar un cambio? *(botones)*
5. ¿Hablamos 15 minutos con la TL? *(botones)*

La 3 es la que más vale: con eso la TL sabe por dónde entrar en la llamada.

**Dos reglas que no se tocan:**

- **Kelly no abre la conversación de reclutamiento.** Entra cuando el candidato
  ya ha respondido y ha consentido. El primero sigue siendo la llamada.
- **Kelly no pregunta por el split.** El mensaje del día 28 promete *«la rellenas
  tú, no me mandas ningún número»*. Si luego el bot le pregunta cuánto se queda,
  se rompe la promesa y con ella la confianza. Solo necesitamos saber si produce
  y qué le duele.

Kelly devuelve las respuestas llamando a `recRecibirCualificacion(payload)` con
un JSON. El sistema recalcula el score con los datos reales, guarda el dolor y
el motivador en la ficha, marca el consentimiento, decide el estado según lo que
haya contestado a la última pregunta y lo registra todo en `Rec_Toques` para la
auditoría. Si el candidato está en la lista de supresión, no guarda nada.

Menú → Kelly/KWSync → *Especificación de cualificación para Kelly* saca la hoja
con las preguntas en ES y EN, los botones y el formato del JSON.

### El hueco que había, y que ya está cerrado

Marbella es pequeña: **un agente de la competencia puede ser además cliente
vuestro**. Si le dijo BAJA a Kelly como cliente y el reclutamiento le seguía
escribiendo como candidato, estabais incumpliendo su derecho de oposición. El
«no» es de la persona, no del canal.

Menú → Kelly/KWSync → *Importar bajas de Kelly* lee la lista de supresión de
KWSync y para en seco a esos candidatos. Rellena `KWSYNC_SHEET_ID` y
`KWSYNC_HOJA_BAJAS` en `Rec_Config` y la automatización diaria lo hace solo,
**antes** de generar los toques del día.

Es de **solo lectura** sobre KWSync: este módulo no escribe nunca en vuestro
sistema de leads en producción. Para el sentido contrario hay un export CSV que
cargáis vosotros.

---

## 10. La calculadora de ingresos (día 28)

`calculadora_ingresos_agente.xlsx` es el activo del paso que más convierte de
toda la secuencia. El agente mete cuatro datos y ve lo que se habría quedado
con cada modelo.

**Tres pestañas:**

| Pestaña | Quién la toca |
|---|---|
| `Instrucciones` | Nadie. Explica el uso y la leyenda de colores |
| `Calculadora` | **El agente**: cuatro celdas amarillas |
| `Parametros_MC` | **Tú, antes de enviarlo**: cinco celdas amarillas |

**Antes de mandarla, rellena `Parametros_MC`:** reparto del agente, tope anual
de aportación, royalty, tope de royalty y cuotas fijas. **Vienen vacías a
propósito** — no me invento las cifras de tu Market Center, y cada MC tiene
las suyas. Si te las dejas vacías, la hoja muestra un aviso en rojo y los
resultados no son válidos.

Puedes ocultar `Parametros_MC` (clic derecho en la pestaña → Ocultar) para que
el agente vea solo su calculadora.

**Por qué funciona:** no le pides sus números. Le das la herramienta y los mete
él. Eso elimina la resistencia de "no te voy a contar lo que gano", y el que
hace el cálculo saca su propia conclusión, que es la única que le mueve.

Lo que más impacta no es el reparto: es la fila del **tope**. Hay un cálculo
que dice cuántas operaciones le hacen falta para alcanzarlo, y una tabla de
escenarios que enseña cómo se abre la diferencia cuanto más produce.

El cálculo está verificado con 50 comprobaciones contra el resultado hecho a
mano, incluidos los casos límite (producción que no llega al tope y hoja vacía).

---

## 11. Seguridad: hay una clave expuesta

El archivo `gs`, línea ~3565, tiene la clave de Gemini escrita en el código:

```javascript
const GEMINI_API_KEY = 'AIzaSyC...';  // ← la clave real está en el archivo, redactada aquí
```

Está en el historial público de este repositorio. Cualquiera que lo vea puede
consumir tu cuota. **Borrar la línea no basta: el historial de Git la conserva.**

Qué hacer, en este orden:

1. Menú **🎯 Reclutamiento → Migrar la clave de Gemini expuesta**.
2. Entra en [aistudio.google.com](https://aistudio.google.com/apikey) → API Keys
   → **borra esa clave**.
3. Crea una clave nueva.
4. Menú **→ Configurar claves de API** → pega la nueva.
5. En el archivo `gs`, **borra** la línea `const GEMINI_API_KEY = ...`.
6. Sustituye las llamadas a `llamarGemini()` por `recLlamarGemini()`, que ya
   lee de `PropertiesService`.

El módulo de reclutamiento no guarda ninguna clave en el código.

---

## 12. Hojas que crea el sistema

| Hoja | Para qué |
|---|---|
| `Rec_Candidatos` | La base. 36 columnas: contacto, producción, score, estado, consentimiento, base legal |
| `Rec_Agencias` | Censo de agencias con reputación y si ya se rastreó su web |
| `Rec_Toques` | Auditoría de cada contacto. Es tu defensa ante una reclamación |
| `Rec_SmartPlan` | Las cadencias y los textos. **Edita aquí**, no en el código |
| `Rec_Entrevistas` | Embudo de entrevistas y datos del Career Visioning |
| `Rec_Supresion` | Lista de bajas. Bloquea todos los canales |
| `Rec_Config` | Parámetros sin tocar código |
| `Rec_RGPD` | Registro de actividad y ponderación de interés legítimo |
| `Rec_Busquedas_LinkedIn` | Las 8 cadenas de búsqueda, listas para pegar |
| `Rec_Cola_Kelly` | Toques entregados a Kelly, con plantilla y variables |
| `Rec_Plantillas_Kelly` | Las plantillas de Meta que faltan para reclutamiento |
| `Rec_Cualificacion_Kelly` | Las 5 preguntas de cualificación y el formato del JSON |

---

## 13. Pruebas

`test_reclutamiento.js` carga el módulo en Node con los servicios de Google
simulados y valida 92 casos: normalización de teléfonos (incluidos los
británicos de los compradores UK), deduplicación, parser de `robots.txt`,
clasificación de perfiles, scoring, plantillas, enlaces de WhatsApp y
supresión, además del parser de perfiles del X-ray de Google el mapeo de plantillas de Kelly y el parser de rangos de botón (incluidos los separadores de miles en inglés).

```bash
node test_reclutamiento.js
```

La calculadora tiene su propia verificación: 50 comprobaciones del cálculo
contra el resultado a mano, con casos límite incluidos.

Ejecútalo siempre que toques los pesos del scoring o la normalización.

---

## Fuentes consultadas

- [Ley 5/2025, de Vivienda de Andalucía (BOJA)](https://www.juntadeandalucia.es/boja/2025/247/1) · [versión BOE](https://www.boe.es/eli/es-an/l/2025/12/16/5)
- [Registro de Agentes Inmobiliarios de Andalucía — qué exige (EANE)](https://www.eane.es/blog/registro-agentes-inmobiliarios-andalucia/) · [estado de desarrollo (Hipotea)](https://hipotea.com/el-registro-obligatorio-de-agentes-inmobiliarios-en-andalucia-ya-es-una-realidad-que-esta-en-vigor-y-como-prepararse-con-tiempo/)
- [WhatsApp Business Platform — obtención de opt-in (Meta)](https://developers.facebook.com/documentation/business-messaging/whatsapp/getting-opt-in) · [Política de mensajería](https://learn.rasayel.io/en/books/whatsapp/whatsapp-business/whatsapp-business-messaging-policy)
- [Límites de WhatsApp a mensajes sin respuesta (TechCrunch, oct. 2025)](https://techcrunch.com/2025/10/17/whatsapp-will-curb-the-number-of-messages-people-and-businesses-can-send-without-a-response)
- [Sanción de la AEPD por spam a través de LinkedIn](https://autelsi.es/observatorioprivacidad/archivos/191788) · [¿Es legal usar CV extraídos de LinkedIn?](https://auratechlegal.es/utilizar-cv-candidatos-linkedin/)
- [Datos de contacto profesionales y art. 19 LOPDGDD (Prodat)](https://www.prodat.es/blog/los-datos-de-contacto-profesionales-y-su-regulacion-en-el-reglamento-europeo-de-proteccion-de-datos/) · [La LSSI no desplaza al RGPD](https://jorgegarciaherrero.com/la-lssi-no-desplaza-al-rgpd-el-rgpd-abraza-a-la-lssi/)
- [Guía de la AEPD sobre protección de datos y relaciones laborales](https://www.aepd.es/prensa-y-comunicacion/notas-de-prensa/aepd-publica-guia-pd-y-relaciones-laborales)
- [CommandMC — SmartPlans de Recruits](https://documentation.kw.com/docs-command/html/kellercloud/commandmc/recruits/manage-recruits/add-recruit-smartplan.html) · [Crear un SmartPlan](https://documentation.kw.com/docs-command/html/kellercloud/command/smartplans/create-custom-smartplan.html)
- [Command no tiene WhatsApp nativo — petición abierta en el foro de ideas de KW](https://ideas.kw.com/forums/959210-command/suggestions/50787875-proposal-for-native-whatsapp-integration-in-kw-com)
- [KW abre Command a desarrolladores externos, feb. 2026 (Inman)](https://www.inman.com/2026/02/23/keller-williams-opens-command-platform-to-3rd-party-developers/) · [nota oficial](https://www.businesswire.com/news/home/20260223307158/en/Keller-Williams-Opens-Command-to-Power-Agent-Choice-and-Best-in-Class-Integrations)
- [KWIQ, el asistente de IA de KW (HousingWire)](https://www.housingwire.com/articles/keller-williams-launches-ai-powered-real-estate-assistant/)
- [Google Places API (New) — Text Search](https://developers.google.com/maps/documentation/places/web-service/text-search)
- [Autónomos inmobiliarios en Andalucía (El Confidencial Digital)](https://www.elconfidencialdigital.com/articulo/legal/andalucia-suma-12500-autonomos-inmobiliarios-cinco-anos/202610071637561491500.html)
