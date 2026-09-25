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

## 2. Paso 0 (innegociable): KWY tiene que estar en Git

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

## 3. Las tres formas de usar Claude Code (elegir una)

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

## 4. Las cuatro cosas que de verdad multiplican la velocidad en KWY

Instalar Claude Code quita el copia-pega. Estas cuatro quitan el resto de la fricción.

### 4.1. Un fichero `CLAUDE.md` en la raíz del proyecto

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

### 4.2. Firebase CLI instalado

Para que Claude Code pueda cerrar el ciclo completo él solo — editar, probar en local y
desplegar — sin que nadie copie comandos:

```powershell
npm install -g firebase-tools
firebase login
```

Con esto, Claude puede ejecutar `firebase emulators:start` para probar los cambios en local,
leer los logs de las Cloud Functions y desplegar cuando se le dice. Hoy esos comandos se están
copiando y pegando a mano; a partir de aquí, no.

### 4.3. Trabajar en ramas, no sobre producción

Con 700 usuarios, un error desplegado se nota. El flujo sano:

```powershell
git checkout -b arreglo-filtro-videos   # rama nueva
# ... se trabaja con Claude ...
firebase emulators:start                # se prueba en local
git checkout main && git merge arreglo-filtro-videos
firebase deploy                         # solo cuando está comprobado
```

Claude Code gestiona todo esto solo si se le pide; lo que importa es que sea la costumbre.

### 4.4. Comandos propios para las tareas repetitivas

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

## 5. Un aviso sobre el tamaño de los ficheros

Vale la pena mirar esto, porque probablemente sea **parte de la causa** del copia-pega actual.

En el repositorio `codigo` de KWSync hay ficheros de este tamaño:

| Fichero         | Líneas | Tamaño |
|-----------------|--------|--------|
| `html`          | 19.548 | 964 KB |
| `html reducido` | 15.009 | 739 KB |
| `gs`            |  6.123 | 241 KB |

Un fichero HTML de casi 20.000 líneas no cabe cómodamente en una ventana de chat. Eso obliga a ir
pegando fragmentos, y a que la IA trabaje sin ver el conjunto — que es justo la dinámica lenta y
propensa a errores que se quiere eliminar. Si KWY tiene la misma forma (un HTML gigante con todo
el CSS y el JavaScript dentro), pasa lo mismo.

Con Claude Code el problema se alivia mucho, porque lee y edita los ficheros por partes en lugar
de necesitarlos enteros en el contexto. Pero conviene además **partirlos**: separar el CSS a su
fichero, el JavaScript por módulos, el HTML en componentes. Es una tarea que se le puede encargar
a Claude Code directamente, con una condición: hacerlo en una rama, con el emulador delante, y
verificando que todo sigue funcionando antes de desplegar.

---

## 6. Plan de arranque: unos 45 minutos

| # | Tarea | Tiempo |
|---|-------|--------|
| 1 | Instalar la extensión de Claude Code en VS Code e iniciar sesión | 5 min |
| 2 | Poner KWY bajo Git y subirlo a un repositorio privado de GitHub (revisando el `.gitignore`) | 15 min |
| 3 | Instalar Firebase CLI (`npm install -g firebase-tools`) y hacer `firebase login` | 5 min |
| 4 | Crear el `CLAUDE.md` con `/init` y repasarlo a mano | 15 min |
| 5 | Poner el modo de permisos en *Auto* y hacer un primer cambio pequeño de prueba | 5 min |

A partir del minuto 45, el ciclo de trabajo es: **describir lo que se quiere → revisar el diff →
aceptar**. Cero copia-pega.

---

## 7. Resumen para quien no quiera leer lo anterior

1. Instalar la **extensión de Claude Code en VS Code**. Esto por sí solo elimina el copia-pega.
2. Poner **KWY en Git/GitHub**. Sin esto no hay red de seguridad para una plataforma de 700 usuarios.
3. Escribir un **`CLAUDE.md`** con el contexto de la plataforma, para no reexplicarlo cada día.
4. Instalar el **Firebase CLI**, para que Claude pueda probar y desplegar sin intermediarios.
5. Usar **modo Plan y ramas** para cualquier cambio serio en producción.

---

## Enlaces

- Instalación y requisitos: https://code.claude.com/docs/en/setup
- Extensión de VS Code: https://code.claude.com/docs/en/vs-code
- Primeros pasos: https://code.claude.com/docs/en/quickstart
- Claude Code en la web: https://claude.ai/code
