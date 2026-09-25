# Cómo dejar de copiar y pegar: Claude Code para mantener KWY

> Guía práctica para el mantenimiento diario de la plataforma KWY (Firebase).
> Objetivo: eliminar el ciclo "copio de VS Code → pego en Claude → copio la respuesta → pego en VS Code".

---

## 1. El diagnóstico, en una frase

El problema **no es Claude ni es VS Code**: es que hoy están desconectados. Claude actúa como
cerebro en una pestaña del navegador y VS Code como manos en otra ventana, y el puente entre
ambos es una persona haciendo Ctrl+C / Ctrl+V. Ese puente es el cuello de botella.

**Claude Code es exactamente la herramienta que elimina ese puente**: es Claude ejecutándose
*dentro* del proyecto, con acceso directo a los ficheros, a la terminal (PowerShell) y a Git.
Lee los ficheros él mismo, los edita él mismo y ejecuta los comandos él mismo. No hay nada que
pegar en ninguna dirección.

Una aclaración de nomenclatura, porque genera confusión: el editor es **Visual Studio Code**
(VS Code). *Visual Basic* es otra cosa, un lenguaje antiguo de Microsoft. Y **PowerShell** es
la terminal de Windows que VS Code lleva integrada abajo: no es un producto aparte al que haya
que ir y volver, es una pestaña más dentro del mismo editor.

---

## 2. Por qué Claude le dijo "ya no podemos seguir así"

No puedo saber el momento exacto sin ver esa conversación, pero las causas realistas son estas
cuatro, y casi siempre es una combinación:

1. **El proyecto se hizo más grande que una ventana de chat.** Una web es muchos ficheros. Un chat
   solo puede llevar uno pegado por mensaje. En cuanto KWY pasó de un puñado de ficheros — o uno
   solo se hizo demasiado grande para pegarlo entero — el método dejó de funcionar.
2. **El chat no puede tocar el servidor.** kwwise.es está desplegado en algún sitio. Una
   conversación de chat no puede subir ficheros, ejecutar un despliegue, leer los logs ni
   comprobar si la web sigue funcionando. En cuanto el trabajo pasó de "escríbeme esta función" a
   "arregla lo que está roto en producción", el chat chocó contra un muro.
3. **Los límites de longitud de la conversación.** Los chats largos con ficheros grandes pegados
   se agotan y hay que empezar de cero, perdiendo todo el contexto cada vez.
4. **Pérdida de fidelidad.** Al pegar un fichero enorme y pedir que lo devuelva entero, se pierden
   trozos por el camino. Es un método que introduce errores en silencio.

**Aquí está la clave: el diagnóstico de Claude era correcto, pero la solución que dio era solo la
mitad.** "Deja de pasar el código por el chat, trabaja sobre los ficheros reales con un editor y
una terminal" es exactamente lo que había que hacer. Lo que faltó fue la segunda mitad: *"…y usa
Claude Code, que trabaja sobre esos ficheros él mismo"*.

Es decir: **no se ha ido hacia atrás, se ha quedado a un paso de terminar la mudanza.** Hizo la
parte incómoda (montar el entorno local) sin la parte que la hace rápida. Por eso ahora está
haciendo a mano el trabajo que la herramienta hace sola.

---

## 3. Test de 30 segundos: ¿qué es eso que tiene abierto dentro de VS Code?

Que Claude aparezca dentro de VS Code **no significa que sea Claude Code**. Hay tres cosas
distintas que se ven parecidas, y solo una toca los ficheros.

**Cuidado con un espejismo muy fácil de creer:** que la lista de ficheros se vea en la columna
izquierda no demuestra nada. Esa columna es el explorador de VS Code, y es independiente del panel
de Claude. Que compartan ventana no los conecta. Un chat embebido en un panel **no ve** lo que
muestra el explorador de al lado, igual que un navegador abierto en la otra mitad de la pantalla
tampoco lo ve. Esa es exactamente la ilusión que hace pensar que "Claude está dentro pero no
funciona solo".

La prueba tiene que ser sobre el **contenido** de un fichero, no sobre su nombre — un chat puede
adivinar que existe un `index.html` y sonar convincente:

> Con la carpeta del proyecto abierta, elegir en la columna izquierda un fichero concreto y
> escribirle al panel de Claude:
>
> **"Abre el fichero `<nombre exacto>`, dime cuántas líneas tiene y resúmeme qué hace."**

Según lo que conteste:

| Respuesta | Qué es | Qué hacer |
|---|---|---|
| Lo abre y da datos reales y verificables | **Es Claude Code y funciona.** El problema es de configuración | Ir al indicador de modo, abajo en la caja de texto, y ponerlo en **Auto**. Si está en *Plan* o *Manual*, Claude solo describe o pide permiso a cada paso — y eso se parece mucho a "me dice lo que tengo que hacer" |
| Dice que no puede acceder a tus ficheros, pide que le pegues el contenido, o responde en genérico sin datos concretos | **Es el chat, metido dentro de VS Code** (un panel de navegador o una extensión de terceros). Parece integrado pero tiene cero acceso al disco. Esta es la trampa | Desinstalarlo e instalar la extensión oficial |
| Sale una pantalla de *Sign in*, o "Not logged in · Please run /login" | Es la extensión oficial, **pero sin sesión iniciada** | Iniciar sesión. Requiere plan de pago (Pro, Max, Team o Enterprise); el plan gratuito no incluye Claude Code |
| No responde nada útil y no hay carpeta abierta | Claude Code **sin proyecto**. Sin carpeta abierta no tiene sobre qué trabajar | `Archivo → Abrir carpeta` y seleccionar la carpeta de kwwise.es |

Para identificar la extensión buena: se llama exactamente **"Claude Code"**, el editor es
**Anthropic**, y el icono es una chispa (✻). Cualquier otra cosa del marketplace con "Claude" en
el nombre es de terceros y no tiene acceso a los ficheros.

---

## 4. Qué escribirle exactamente a su Claude

Esto se puede copiar y pegar tal cual en el panel de Claude:

```text
Para. Vamos a cambiar de método.

Estoy usando Claude Code dentro de VS Code, con la carpeta del proyecto de
kwwise.es abierta. Quiero que trabajes directamente sobre los ficheros: que
los leas tú, los edites tú y ejecutes tú los comandos en la terminal.

Deja de darme instrucciones para que yo las copie y las ejecute a mano, y
deja de pedirme que te pegue el contenido de los ficheros. Si necesitas ver
un fichero, ábrelo tú. Si necesitas ejecutar algo, ejecútalo tú.

Para empezar: recorre tú la carpeta, hazme un resumen de la estructura del
proyecto, dime con qué está construida la web y cuántas líneas tiene cada
fichero principal.
```

La última frase es la prueba de fuego, y por eso pide **cifras**: los nombres de fichero se pueden
adivinar, el número de líneas no. **Si vuelve con datos concretos y comprobables, ya está
resuelto** y se acabó el copia-pega. Si pide que le peguen el contenido, o contesta en genérico,
no es Claude Code: hay que volver al test del apartado anterior.

Y si contesta bien pero sigue limitándose a describir en vez de tocar nada, entonces está en modo
*Plan*. Segundo mensaje:

```text
Deja el modo plan y haz los cambios directamente sobre los ficheros.
```

(O cambiarlo a mano en el indicador de modo, abajo en la caja de texto.)

---

## 5. Paso 0 (innegociable): KWY tiene que estar en Git

Antes de cualquier otra cosa. Una plataforma con ~700 agentes y un repositorio de vídeos por
roles no se puede mantener sin control de versiones:

- Sin Git no hay forma de **deshacer** un cambio que rompe producción.
- Sin Git no hay historial de **qué se cambió, cuándo y por qué**.
- Sin Git no se puede usar Claude Code desde la web ni desde el móvil.
- Sin Git, cada sesión de IA empieza a ciegas sobre el estado real del código.

En la cuenta de GitHub conectada actualmente solo aparece el repositorio `codigo`. Si KWY vive
únicamente en una carpeta local del portátil y se sube a Firebase Hosting con `firebase deploy`,
está a un disco duro roto de desaparecer.

Desde PowerShell, dentro de la carpeta de KWY:

```powershell
git init
git add .
git commit -m "Estado inicial de KWY"
```

Y después crear el repositorio en GitHub (privado) y subirlo. Si se prefiere, este paso se lo
puede pedir directamente a Claude Code una vez instalado: sabe hacerlo solo.

> **Importante antes del primer commit:** revisar que no se suben claves. En un proyecto Firebase,
> los ficheros de credenciales de servicio (`serviceAccountKey.json`, `*-firebase-adminsdk-*.json`)
> y los `.env` deben ir en `.gitignore`. La configuración pública del cliente (`apiKey` del
> `firebaseConfig`) sí puede ir en el repo: es pública por diseño, lo que protege los datos son
> las reglas de seguridad de Firestore.

---

## 6. Las tres formas de usar Claude Code (elegir una)

### Opción A — Extensión de VS Code · **la recomendada**

Es la que menos cambia la forma de trabajar de alguien que ya vive en VS Code. Claude aparece
como un panel lateral en el mismo editor: se le habla en lenguaje natural, propone los cambios
como un **diff lado a lado** (antes / después) y se aceptan o rechazan con un clic.

**Requisitos:** VS Code 1.94 o superior + una suscripción de pago a Claude (Pro, Max, Team o
Enterprise). El plan gratuito de claude.ai no incluye Claude Code.

**Instalación:**
1. En VS Code, `Ctrl+Shift+X` para abrir Extensiones.
2. Buscar **"Claude Code"** (editor: Anthropic) e **Instalar**.
3. Abrir el panel con el icono de la chispa (arriba a la derecha del editor, o en la barra
   lateral izquierda).
4. Pulsar **Sign in** y autorizar en el navegador con la cuenta de Claude.

A partir de ahí: abrir la carpeta de KWY en VS Code (`Archivo → Abrir carpeta`) y escribir en el
panel lo que se quiere hacer. Nada más.

**Dos ajustes que cambian la experiencia:**
- **Modo de permisos** (abajo en la caja de texto): en modo *Auto* Claude edita sin preguntar a
  cada paso. Para el trabajo del día a día es lo que evita la fatiga de confirmar todo.
- **Modo Plan** (`/plan`): para cambios grandes en una plataforma en producción. Claude describe
  primero qué va a hacer, se revisa, se corrige, y solo entonces toca el código. En una plataforma
  con 700 usuarios esto no es opcional para cambios de calado.

### Opción B — Claude Code en la terminal integrada

Misma potencia, sin extensión. Se ejecuta en la pestaña de PowerShell que VS Code ya tiene abajo,
así que sigue siendo *una sola ventana*.

```powershell
# 1. Instalar (PowerShell, no hace falta ser administrador)
irm https://claude.ai/install.ps1 | iex

# 2. Comprobar que ha ido bien
claude --version

# 3. Ir a la carpeta del proyecto y arrancar
cd C:\ruta\a\kwy
claude
```

Se recomienda instalar también [Git para Windows](https://git-scm.com/downloads/win): sin él,
Claude Code usa PowerShell para ejecutar comandos; con él dispone además de Bash, que le da más
capacidad.

### Opción C — Claude Code en la web y en el móvil

En [claude.ai/code](https://claude.ai/code) se conecta el repositorio de GitHub y Claude trabaja
en un contenedor en la nube: hace los cambios, los sube a una rama y abre una Pull Request.
Requiere haber hecho el **Paso 0**.

Es el complemento ideal, no el sustituto: sirve para lanzar una tarea desde el móvil ("añade el
rol X al filtro de vídeos") y revisar el resultado después desde el ordenador. *Esta misma guía
se ha escrito así.*

---

## 7. Las cuatro cosas que de verdad multiplican la velocidad en KWY

Instalar Claude Code quita el copia-pega. Estas cuatro quitan el resto de la fricción.

### 7.1. Un fichero `CLAUDE.md` en la raíz del proyecto

Es, con diferencia, **el mayor ahorro de tiempo diario**. Claude Code lo lee automáticamente al
arrancar en cada sesión, así que deja de ser necesario reexplicar el contexto de la plataforma
una y otra vez. Un esquema razonable para KWY:

```markdown
# KWY — Plataforma de formación para agentes Keller Williams

## Qué es
Plataforma interna para ~700 agentes. Repositorio de vídeos organizado por roles.

## Stack
- Hosting: Firebase Hosting
- Base de datos: Firestore (colecciones: `usuarios`, `videos`, `roles`, ...)
- Autenticación: Firebase Auth
- Frontend: [describir]

## Estructura
- `public/` — lo que se despliega
- `functions/` — Cloud Functions
- `firestore.rules` — reglas de seguridad

## Roles y permisos
[Qué rol ve qué vídeos]

## Cómo se despliega
`firebase deploy --only hosting`

## Normas
- Nunca desplegar a producción sin probar antes en el emulador.
- Los cambios en `firestore.rules` se revisan siempre a mano.
- [Convenciones de código propias]
```

Se puede generar un primer borrador pidiéndoselo a Claude Code con el comando `/init`, y luego
corregirlo a mano.

### 7.2. La herramienta de despliegue, en la misma máquina

Para que Claude Code pueda cerrar el ciclo completo él solo — editar, probar en local y
desplegar — sin que nadie copie comandos.

**Si kwwise.es está en Firebase Hosting:**

```powershell
npm install -g firebase-tools
firebase login
```

Con esto, Claude puede ejecutar `firebase emulators:start` para probar los cambios en local,
leer los logs de las Cloud Functions y desplegar cuando se le dice. Hoy esos comandos se están
copiando y pegando a mano; a partir de aquí, no.

**Si está en un hosting clásico (FTP, cPanel, Plesk, un VPS):** el principio es el mismo, cambia
la herramienta. Lo que hay que conseguir es que el despliegue sea *un comando ejecutable desde la
terminal*, no una subida manual arrastrando ficheros. Un script `deploy.ps1` con el `rsync`, el
`scp` o el cliente FTP correspondiente es suficiente. A partir de ahí Claude Code lo ejecuta él.

La forma rápida de resolverlo es preguntárselo a él directamente:

```text
Explícame cómo se despliega este proyecto ahora mismo, y escríbeme un script
que lo haga en un solo comando. Luego lo añadimos al CLAUDE.md.
```

### 7.3. Trabajar en ramas, no sobre producción

Con 700 usuarios, un error desplegado se nota. El flujo sano:

```powershell
git checkout -b arreglo-filtro-videos   # rama nueva
# ... se trabaja con Claude ...
firebase emulators:start                # se prueba en local
git checkout main && git merge arreglo-filtro-videos
firebase deploy                         # solo cuando está comprobado
```

Claude Code gestiona todo esto solo si se le pide; lo que importa es que sea la costumbre.

### 7.4. Comandos propios para las tareas repetitivas

Todo lo que se haga más de tres veces se puede convertir en un comando de una línea. Se crean
como ficheros `.md` en `.claude/commands/` dentro del proyecto. Por ejemplo,
`.claude/commands/nuevo-video.md`:

```markdown
Añade un vídeo nuevo al catálogo de KWY:
1. Pídeme título, URL, rol o roles que pueden verlo y categoría.
2. Añade el documento a la colección `videos` de Firestore.
3. Comprueba que los permisos por rol son coherentes con los vídeos ya existentes.
4. Enséñame el cambio antes de desplegar.
```

A partir de ahí, escribir `/nuevo-video` en Claude Code dispara toda la secuencia.

---

## 8. Un aviso sobre el tamaño de los ficheros

Este apartado puede que no aplique a KWY: si en el explorador de VS Code se ven muchos ficheros,
el proyecto ya está repartido y no hay nada que partir. Aun así conviene comprobar que ninguno se
ha ido de las manos, porque es un problema silencioso.

En el repositorio `codigo` de KWSync hay ficheros de este tamaño:

| Fichero         | Líneas | Tamaño |
|-----------------|--------|--------|
| `html`          | 19.548 | 964 KB |
| `html reducido` | 15.009 | 739 KB |
| `gs`            |  6.123 | 241 KB |

Un fichero HTML de casi 20.000 líneas no cabe cómodamente en una ventana de chat. Eso obliga a ir
pegando fragmentos, y a que la IA trabaje sin ver el conjunto — que es justo la dinámica lenta y
propensa a errores que se quiere eliminar.

Para salir de dudas en KWY, basta con preguntárselo a Claude Code una vez conectado: *"dime los
diez ficheros más largos del proyecto y cuántas líneas tiene cada uno"*. Por encima de unas 2.000
líneas en un solo fichero, conviene plantearse dividirlo.

Con Claude Code el problema se alivia mucho, porque lee y edita los ficheros por partes en lugar
de necesitarlos enteros en el contexto. Pero conviene además **partirlos**: separar el CSS a su
fichero, el JavaScript por módulos, el HTML en componentes. Es una tarea que se le puede encargar
a Claude Code directamente, con una condición: hacerlo en una rama, con el emulador delante, y
verificando que todo sigue funcionando antes de desplegar.

---

## 9. Plan de arranque: unos 45 minutos

| # | Tarea | Tiempo |
|---|-------|--------|
| 1 | Instalar la extensión de Claude Code en VS Code e iniciar sesión | 5 min |
| 2 | Poner KWY bajo Git y subirlo a un repositorio privado de GitHub (revisando el `.gitignore`) | 15 min |
| 3 | Dejar el despliegue como un comando de terminal (Firebase CLI, o un script `deploy`) | 5 min |
| 4 | Crear el `CLAUDE.md` con `/init` y repasarlo a mano | 15 min |
| 5 | Poner el modo de permisos en *Auto* y hacer un primer cambio pequeño de prueba | 5 min |

A partir del minuto 45, el ciclo de trabajo es: **describir lo que se quiere → revisar el diff →
aceptar**. Cero copia-pega.

---

## 10. Resumen para quien no quiera leer lo anterior

1. Instalar la **extensión de Claude Code en VS Code**. Esto por sí solo elimina el copia-pega.
2. Poner **KWY en Git/GitHub**. Sin esto no hay red de seguridad para una plataforma de 700 usuarios.
3. Escribir un **`CLAUDE.md`** con el contexto de la plataforma, para no reexplicarlo cada día.
4. Dejar el **despliegue como un comando de terminal** (Firebase CLI o un script `deploy`), para que Claude pueda probar y desplegar sin intermediarios.
5. Usar **modo Plan y ramas** para cualquier cambio serio en producción.

---

## Enlaces

- Instalación y requisitos: https://code.claude.com/docs/en/setup
- Extensión de VS Code: https://code.claude.com/docs/en/vs-code
- Primeros pasos: https://code.claude.com/docs/en/quickstart
- Claude Code en la web: https://claude.ai/code
