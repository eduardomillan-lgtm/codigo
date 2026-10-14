# Incorporar el motor de reclutamiento a KWSync

Documento para quien construya la integración. Resume qué es el módulo, **qué
NO hay que portar porque KWSync ya lo tiene**, qué sí, y las decisiones que no
conviene deshacer.

---

## 1. Qué es

Un módulo de Apps Script que capta agentes inmobiliarios de fuentes públicas,
los puntúa por probabilidad de incorporación y los lleva por una secuencia de
seguimiento hasta la entrevista con la Team Leader.

Hoy vive sobre el Dashboard Inmobiliario KW (el proyecto con `gs`, `html` y
`html reducido`), en 13 hojas con prefijo `Rec_`.

| | |
|---|---|
| `reclutamiento.gs` | 4.042 líneas · 92 funciones públicas |
| `panel_reclutamiento.html` | Panel diario de la Team Leader |
| `RECLUTAMIENTO.md` | Playbook operativo, 14 secciones |
| Pruebas | 176 comprobaciones en 3 suites, todas en verde |
| Activos | Calculadora de ingresos (.xlsx), semilla de 40 agencias, objetivos |

---

## 2. Lo que NO hay que portar

**Esta es la sección importante.** KWSync ya tiene estas piezas, probadas en
producción. Portar mi versión crearía dos fuentes de verdad, y en tres casos
eso es directamente un incumplimiento.

| Mi pieza | Qué usar de KWSync | Por qué |
|---|---|---|
| `recClaveDedupe_` / `recDeduplicar` | **Su deduplicador** (últimos 9 dígitos + email) | El suyo está probado con volumen real y ya resuelve el caso del Market Center, donde el mismo contacto de dos agentes NO se fusiona |
| `Rec_Consentimientos` | **La página de consentimiento de Kelly** | La de Kelly está en 7 idiomas y guarda el texto exacto, fecha, versión e idioma. La mía es una versión más pobre de lo mismo |
| `Rec_Supresion` | **Su lista de bajas, una sola** | ⚠️ **CRÍTICO.** Dos listas separadas es el agujero que ya detecté y cerré con una sincronización. Al integrar, que haya UNA. Un agente de la competencia puede ser además cliente vuestro: el «no» es de la persona, no del canal |
| `Rec_Toques` como log | **Su registro general** | El suyo ya es buscable por teléfono, referencia y fecha |
| `recEnviarWhatsAppAPI_` | **Kelly** | Ya tiene Cloud API, plantillas aprobadas, ventana de 24 h y topes. Mi implementación quedó deprecada a propósito |
| `panel_reclutamiento.html` | **Una pestaña más en su panel** | Su panel ya tiene 13 pestañas con un patrón establecido. Añadir «Reclutamiento» es más coherente que un panel aparte |
| Aviso diario por email | **Su sistema de alertas** | El suyo ya dice *qué hacer*, no solo qué ha pasado. Las alertas de reclutamiento encajan ahí |
| Copias de seguridad | **Las suyas, diarias** | Nada que hacer |

---

## 3. Lo que sí hay que portar

### 3.1 Conectores de captación — el núcleo

Ninguno existe en KWSync. Son lo que diferencia este módulo.

| Fuente | Qué da | Estado |
|---|---|---|
| **⭐ RealAdvisor** | Agentes **ya ordenados por transacciones de los últimos 24 meses**, con ventas, precio mediano y años de oficio | La mejor. ~47 en Marbella, ~58 en Málaga |
| **Webs de agencias** | Equipo comercial: nombre, cargo, email, teléfono, idiomas. Respeta `robots.txt`, extrae con IA | 1.500-4.000 agentes potenciales |
| **X-ray de LinkedIn vía Google CSE** | Nombre, cargo, agencia, URL. Sin teléfono | 100 consultas/día gratis |
| **Google Places** | Censo de agencias | Opcional, de pago |
| **Pegado de portales** | Producción real por agente | Semiautomático a propósito (ver §5.6) |
| **Semilla de agencias por CSV** | Arranque sin API de Places | 40 agencias ya investigadas |
| **Generador de búsquedas** | 8 segmentos de LinkedIn + X-ray | Documental |

> **Ojo con el solapamiento:** KWSync ya ingiere leads de 8 portales por correo.
> Esos correos llevan el nombre de la agencia y a veces el teléfono del agente
> del anuncio. **Es una fuente de candidatos que ya tenéis dentro** y que mi
> módulo no explota. Merece la pena mirarlo antes de portar el pegado manual.

### 3.2 Scoring de reclutamiento

8 factores con pesos que suman 100: producción (30), zona (15), perfil (15),
idiomas (10), experiencia (10), señales de dolor (10), accesibilidad (5),
red (5). Temperatura A/B/C/D.

No tiene nada que ver con el scoring de un lead de cliente. Es un modelo aparte.

### 3.3 Lógica de segmento → argumento

Lo más valioso del proyecto y lo que no está en ningún sitio más.

| Segmento | Qué NO funciona | Qué sí |
|---|---|---|
| Autónomo con marca propia | **El split** — ya se queda más de lo que le ofreces | Estructura y propiedad de la cartera |
| Red de independientes (SAFTI 70→100%, iad ~87,8%) | **El split y el tope** — ya tienen los dos | Oficina física y equipo: son redes 100% remotas |
| Century 21 / Coldwell (techo ~60%, sin fijo) | — | El split y el tope, directos |
| Agencia local (~40-50%) | — | El split y **la propiedad de la cartera** |
| Franquicia de lujo | Criticar su marca | Comprobar si es socio: si lo es, es aliado, no candidato |
| Cambio de sector | Hablar de comisiones | *«¿De qué vivo mientras aprendo?»* |

Se deduce automáticamente: si el nombre de la persona está dentro del nombre de
la agencia, es su negocio, y la ficha avisa de **no abrir por comisión**.

### 3.4 Smart Plans de reclutamiento

5 cadencias con copys en ES y EN, editables en hoja: `AGENTE_ACTIVO` (15 toques
/90 días), `AUTONOMO`, `CAMBIO_SECTOR`, `NURTURE`, `POST_ENTREVISTA` (incluye la
estructura de Career Visioning).

**No son los SmartPlans de Kelly.** Kelly cualifica compradores; esto es una
secuencia de reclutamiento multicanal con llamada como primer toque.

### 3.5 Resto

- **Embudo de entrada** (`recDoPost`): formulario con consentimiento probado →
  candidato con base jurídica de consentimiento. Podría reimplementarse sobre la
  página de consentimiento de Kelly, que ya existe.
- **Cualificación de candidatos por Kelly**: 5 preguntas, especificación lista.
- **Calculadora de ingresos**: el imán del embudo. 53 fórmulas, 50 comprobaciones.
- **Export a CommandMC Recruits**: módulo distinto de Command Contactos.
- **Capa RGPD**: base legal por ficha, ponderación de interés legítimo, art. 14,
  retención.

---

## 4. Modelo de datos

13 hojas. Las tres que importan:

**`Rec_Candidatos`** — 36 columnas. Las no obvias:

| Columna | Para qué |
|---|---|
| `Inmuebles_Activos` / `Precio_Medio` | Producción. De portal (estimada) o de RealAdvisor/Kelly (real) |
| `Score` / `Temperatura` | **Son cosas distintas.** Score = encaje. Temperatura = orden de llamada. No fusionarlas (ver §5.8) |
| `Plan_Activo` / `Paso_Plan` / `Próximo_Toque` | Estado de la secuencia |
| `Consentimiento_WA` / `Fecha_Consentimiento` | Puerta de entrada a Kelly |
| `Base_Legal` | Interés legítimo o consentimiento, según cómo entró |
| `Fuente` / `URL_Fuente` | **Obligatorio para el RGPD.** De dónde salió cada dato |

**`Rec_Agencias`** — censo con señal de rotación y si ya se rastreó la web. Es
distinta de la cartera de KWSync: son agencias de la competencia, no inmuebles.

**`Rec_SmartPlan`** — las cadencias y los textos. Se edita en hoja, no en código.

---

## 5. Decisiones que no conviene deshacer

Cada una tiene una razón concreta. Están ordenadas por lo que cuesta
equivocarse.

### 5.1 El primer toque es una LLAMADA. Nunca al revés

El art. 21 LSSI exige consentimiento previo para comunicaciones comerciales
electrónicas, y la AEPD ya ha sancionado por guardar datos de un perfil público
de LinkedIn y enviar un email comercial sin consentimiento. La llamada a un
teléfono que el propio agente publica **para recibir llamadas profesionales** es
el primer contacto más defendible. Y convierte más.

### 5.2 Número de WhatsApp SEPARADO para reclutamiento

El de Kelly es el de la oficina y es el que contesta a los leads de cliente en
menos de un minuto. La prospección en frío es el tráfico que genera bloqueos, y
un número con la calificación caída pierde límite de envío. **No se arriesga la
respuesta a leads para ahorrar clics.**

### 5.3 Kelly no abre conversaciones de reclutamiento

Su premisa legal es *«la consulta la hizo él»*. Un agente al que reclutas no ha
preguntado nada. Kelly entra en el toque 3, con consentimiento ya dado.

### 5.4 Kelly no pregunta por el split

El mensaje del día 28 promete literalmente *«la rellenas tú, no me mandas ningún
número»*. Si el bot luego pregunta cuánto se queda, se rompe la promesa.
**Hay una prueba que falla si alguien añade esa pregunta.**

### 5.5 Las plantillas de reclutamiento son categoría MARKETING

Las de Kelly son de servicio. La autorización de Meta va ligada a la finalidad:
reutilizar una de servicio para prospección es motivo de sanción. Y la categoría
la asigna Meta, no vosotros.

### 5.6 Los portales se pegan a mano, no se rastrean

Idealista, Fotocasa y el resto lo prohíben en sus condiciones y en `robots.txt`.
Pegando la página es una persona consultando una web pública, que es para lo que
está publicada.

### 5.7 Nada de scrapers de LinkedIn

Su API de mensajes es solo para partners y **exige que el mensaje vaya ligado a
una acción concreta del usuario**: un envío automático no cuenta. Las
herramientas de terceros envían por tu cuenta conectada, fuera de la plataforma
autorizada. Si restringen la cuenta, la red de la TL no se recupera.

### 5.8 Score y temperatura son cosas distintas

Score mide el encaje. Temperatura manda el orden de llamada. Un entrante va
**siempre en A** aunque su producción puntúe B: ha rellenado un formulario sobre
cuánto gana, está mirando opciones ahora. Un B que levanta la mano vale más que
un A que no te conoce. El score no se infla.

### 5.9 Un dato que no se tiene se deja VACÍO

En la semilla de agencias, 15 de 40 van sin dominio. Un dominio inventado
envenena la base y hace que el rastreador se coma errores en silencio. Lo mismo
con los corchetes `[así]` de los mensajes: si no puedes respaldar una cifra,
bórrala.

### 5.10 Los agentes de otros MC de KW se excluyen

No entran en secuencia. Está en `recAutoAsignarPlanes`.

---

## 6. Bugs ya resueltos — no los reintroduzcáis

Todos encontrados probando, todos con prueba que los cubre.

| | Problema | Solución |
|---|---|---|
| 1 | El archivo `gs` ya tiene **dos** `onOpen()` (líneas 54 y 198); el segundo gana. Un tercero rompe el menú | El módulo **no define** `onOpen` ni `doPost`. Se llama a `recCrearMenu()` / `recDoPost()` desde los existentes |
| 2 | Una hoja nueva trae 26 columnas; 36 cabeceras revientan el `setValues` | `insertColumnsAfter` antes de escribir |
| 3 | `recLeerConfig_` leía la hoja entera en cada llamada, dentro de bucles por candidato: miles de lecturas con 500 fichas | Caché a nivel de módulo + `recRefrescarConfig()` |
| 4 | El objeto del paso del plan se compartía entre candidatos: `canalEfectivo` se quedaba pegado y contaminaba a los siguientes del lote | Copia por candidato con `Object.assign` |
| 5 | Separadores de miles: `€10,000` se leía como 10 € | Se quita el separador solo cuando lo es: `/[.,](?=\d{3}(?!\d))/` |
| 6 | El recálculo del score pisaba la prioridad A del entrante | Suelo de A para entrantes, con nota explicando por qué |
| 7 | `SpreadsheetApp.getUi` no se puede sobrescribir en Apps Script | Flag `silencioso` + objeto de UI muda |
| 8 | Clave de Gemini escrita en el código y subida a git | `PropertiesService` + helper de migración |

---

## 7. Lo que queda abierto

Por orden de lo que bloquea:

1. **Rotar la clave de Gemini.** Sigue en `gs:3565` y en el historial de git.
   **Bloquea la extracción de agentes y RealAdvisor**, que dependen de ella.
2. **Segundo número de WhatsApp** + **4 plantillas MARKETING** aprobadas.
3. **Landing page** del embudo de entrada, con el texto de consentimiento
   validado. Falta la razón social del MC.
4. **Validación del asesor de protección de datos** de la ponderación de interés
   legítimo, y alta de la actividad en el Registro del art. 30 RGPD.
5. **Restringir el acceso a la hoja.** Hoy es una hoja de cálculo con datos
   personales de miles de personas.
6. **Vigilar el Registro de Agentes Inmobiliarios de Andalucía.** Ley 5/2025, en
   vigor desde el 24/01/2026, registro público y obligatorio **todavía no
   operativo**; la Junta tiene hasta el 24/01/2028. El día que abra es la mejor
   fuente que va a existir y sustituye a los pasos de censo.

### Y una decisión de arquitectura pendiente

**KWSync es multioficina y tiene panel central para MAO. Este módulo es de un
solo Market Center.** Al integrarlo hay que decidir el alcance por oficina:
quién ve qué candidatos, si la lista de bajas es global o por oficina (debería
ser global por persona), y si el reclutamiento se ofrece a todas las oficinas o
solo a las que tengan Team Leader.

No está resuelto y es la pregunta más grande de la integración.

---

## 8. Referencia rápida

```
recInicializarTodo()            crea las 13 hojas y siembra los Smart Plans
recCrearMenu()                  llamar desde el onOpen existente
recDoPost(e)                    llamar desde un doPost nuevo

CAPTACIÓN
recImportarAgencias(csv)        semilla de agencias
recImportarGooglePlaces()       censo por API
recRastrearWebsAgencias()       extrae equipos, por lotes de 25
recImportarRealAdvisor()        ⭐ agentes ya ordenados por producción
recBuscarXRayGoogle()           perfiles de LinkedIn vía Google CSE
recImportarTexto(c,f,modo)      CSV o texto libre interpretado por IA
recGenerarBooleanLinkedIn()     8 segmentos de búsqueda

MOTOR
recRecalcularScores(sil)        puntúa y ordena
recAutoAsignarPlanes()          inscribe por perfil
recGenerarToques_(sil)          núcleo: genera la cola del día
recMarcarResultadoToque(...)    mueve el pipeline
recRutinaDiaria()               trigger de las 7:00

KELLY
recSincronizarSupresionKWSync() importa sus bajas (solo lectura)
recRecibirCualificacion(p)      entrada de datos de Kelly
recEspecificacionKelly()        las 5 preguntas
recVerPlantillasKelly()         las 4 plantillas MARKETING que faltan

RGPD
recRegistrarOptOut(id,m,n)      baja en todos los canales
recPurgarRetencion()            purga a los 365 días
recTextoConsentimiento(idioma)  texto en ES y EN
```

Pruebas: `node test_reclutamiento.js` (109) · `test_agencias.js` (22) ·
`test_formulario.js` (45).
