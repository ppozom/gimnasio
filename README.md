# Set y serie

Aplicación web personal para registrar entrenamientos de gimnasio serie por serie.
Está hecha con **Google Apps Script** (`Code.gs`) y una interfaz de una sola página
(`Index.html`). Usa como base de datos la planilla de Google **rutina_gimnasio**.

Se publica como aplicación web («Ejecutar como: yo», «Acceso: solo yo») y se usa
desde el navegador del teléfono, en el gimnasio.

## Funciones

- Elegir el día de entrenamiento y ver sus ejercicios con series, repeticiones objetivo, nota, descanso y claves de técnica.
- Registro serie por serie, con peso y repeticiones propios en cada serie. Se pueden añadir y quitar series.
- Tres estados por serie: sin hacer → hecha → hecha pero interrumpida.
- Precarga de lo hecho en el último guardado de cada ejercicio, aunque haya sido en otro día.
- Aviso **«sube 2,5–5 kg»** (o «5–10 lb») cuando la última vez llegaste al tope del rango en todas las series.
- **Superseries:** los ejercicios seguidos con la misma letra en la columna Superserie se muestran juntos; al marcar una serie de uno se pasa al siguiente sin descanso, y el descanso parte al terminar la vuelta.
- **Kilos o libras** por ejercicio: el botón `kg`/`lb` de cada serie cambia la unidad. Se anota lo que marca la máquina, se precarga igual la próxima vez y se muestra la equivalencia en kilos.
- Botones por ejercicio: **Técnica** (claves y video), **Comentario** e **Historial** (últimas 6 sesiones).
- **Comentario vigente:** lo que escribes en un ejercicio se ve en las sesiones siguientes hasta que lo cambies o lo borres. La Observación de Registro guarda solo el día en que cambió.
- Cambio entre el ejercicio original y sus alternativas, cada una con sus propios datos.
- **Descanso** editable por ejercicio:
  - parte solo al marcar una serie (se puede desactivar);
  - pitidos cortos cuando faltan 10 s y alarma de unos 3 s al terminar, aunque la pantalla esté bloqueada;
  - franja inferior con la cuenta regresiva y los botones −15, +15 y detener.
- Cronómetro de la sesión. Si la primera serie se marca más de 15 min después de abrir el día, el cronómetro parte en ese momento.
- Borrador automático en el teléfono (`localStorage`): si la pestaña se recarga, se recupera lo anotado.
- Guardar pide confirmación. Si se reenvía la misma sesión porque la respuesta no llegó (sin señal), no se duplica.
- Tres temas de color: negro con verde, con celeste y con naranja.

## Estructura de la planilla

| Hoja | Columnas |
|---|---|
| **Rutina** | Día · id · Orden · Ejercicio · Series · Reps objetivo · Nota · Video/Foto (URL) · tiempo_descanso · Técnica 1..5 · Superserie |
| **Alternativas** | id · Ejercicio · Series · Reps objetivo · Nota · Video/Foto (URL) · tiempo_descanso · Técnica 1..5 |
| **Registro** | Fecha · Hora · Día · Ejercicio · Serie · Peso (kg) · Reps · Descanso (s) · Observación · Estado · id · Unidad |
| **Sesiones** | Fecha · Día · Inicio · Término · Duración (min) · Ejercicios · Series |
| **Comentarios** | Ejercicio · Comentario · Fecha |

- En **Alternativas** puede haber varias filas por `id`. Si se dejan en blanco Series, Reps o tiempo_descanso, se heredan del ejercicio original.
- En **Rutina**, `Superserie` es una letra (A, B…). Dos o más ejercicios seguidos del mismo día con la misma letra forman una superserie; una letra sola no hace nada.
- En **Registro**, `Estado` es `completa` o `interrumpida`, y `Unidad` es `kg` o `lb` (vacía = kg; así quedan las series anteriores). `Peso (kg)` guarda el número tal como se anotó, en la unidad de su fila.
- En **Comentarios** hay una fila por ejercicio, con su nombre exacto. Un comentario vacío significa que lo borraste.
- `Reps objetivo` y `tiempo_descanso` van con formato de texto plano. `tiempo_descanso` acepta `180` (segundos), `3` (10 o menos se toma como minutos) y `2:30`.

## Restricciones de la plataforma

1. La app corre dentro de un iframe de Google en otro dominio. Ahí no funcionan Wake Lock ni las notificaciones del navegador.
2. Web Audio se suspende con la pantalla bloqueada. Por eso el aviso se arma como un WAV completo (casi silencio durante el descanso y la alarma al final) que se reproduce en un `<audio>`.
3. `localStorage` sí funciona en el iframe: ahí se guardan el tema, la sesión en curso y el borrador.
4. Sheets convierte `5-8` en fecha, `2:00` en hora y `22.5` en la fecha 22-may. El código tolera esas conversiones y guarda el peso como número.

## Despliegue

1. En el editor de Apps Script de la planilla, pega `src/Code.gs` y `src/Index.html`. El archivo HTML debe llamarse `Index`.
2. Si la versión trae cambios de estructura, ejecuta la función de migración que corresponda (ver abajo).
3. Ve a Implementar → Administrar implementaciones → editar → **Nueva versión**. Así se mantiene la misma URL.

## Funciones de mantenimiento (se ejecutan desde el editor)

| Función | Qué hace | Cambia datos |
|---|---|---|
| `configurarHojas()` | Instalación nueva: crea las hojas con la rutina base | Solo en una planilla vacía |
| `migrarAlNuevoFormato()` | Pasa del formato antiguo (columna de sustitutos) al actual | Rutina y Alternativas |
| `migrarSuperseriesLibrasYComentarios()` | Agrega la columna Superserie (con las letras que ya dicen las notas), la cabecera Unidad en Registro y la hoja Comentarios (con la última observación de cada ejercicio) | Rutina (columna nueva) y hoja nueva. Se puede ejecutar dos veces |
| `migrarEjerciciosNuevos()` | Agrega hip thrust, plancha Copenhague, abducción y alternativas nuevas, con su técnica; completa la técnica que falte | Rutina y Alternativas. Se puede ejecutar dos veces |
| `revisarDatos()` | Informa incoherencias en el registro de ejecución | No |
| `repararPesosConFecha()` | Opcional: devuelve a número los pesos que Sheets guardó como fecha | Solo esas celdas de Registro |

Ninguna migración borra historial ni toca Registro o Sesiones, salvo lo indicado en la tabla.

## Enlazar con GitHub

```bash
git remote add origin https://github.com/<usuario>/rutina_gimnasio.git
```

```bash
git push -u origin main
```

Para sincronizar con Apps Script desde la terminal se puede usar [clasp](https://github.com/google/clasp) con `rootDir` apuntando a `src/`. El archivo `.clasp.json` contiene el id del script y está excluido del repositorio.

## Historial de cambios

### 2026-10 (2)
- **Nuevo:** superseries, botón kg/lb por ejercicio y comentario vigente por ejercicio, como en la versión nueva (Next.js + Supabase).
- **Nuevo:** `migrarSuperseriesLibrasYComentarios()`; `revisarDatos()` también revisa letras de superserie, unidades y comentarios.
- **Corregido:** una observación como «5-8» ya no se convierte en fecha al guardarse.

### 2026-10
- **Corregido:** la alarma de fin de descanso no sonaba. Al llegar a cero, el código pausaba el audio justo antes de los pitidos.
- **Corregido:** si había dos guardados el mismo día, la precarga sumaba sus series (Remo T aparecía con 8).
- **Corregido:** los pesos con decimales se guardaban como fecha y la precarga los dejaba vacíos.
- **Nuevo:** descanso automático al marcar una serie, pre-aviso a los 10 s, alarma más audible y franja de descanso con −15/+15.
- **Nuevo:** historial por ejercicio, aviso de progresión, confirmación al guardar y guardado sin duplicados.
- **Nuevo:** hip thrust, plancha Copenhague, abducción de cadera, rotación externa, saltos laterales y la alternativa recuperada «Jalón con agarre supino».
- **Nuevo:** `revisarDatos()` y `repararPesosConFecha()`.
- **Quitado:** el intento de mantener la pantalla encendida y las notificaciones, que no funcionan en Apps Script.
